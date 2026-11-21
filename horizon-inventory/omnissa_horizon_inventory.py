#!/usr/bin/env python3
"""
omnissa_horizon_inventory.py

Collects Omnissa Horizon 8 inventory through the Connection Server REST API
and writes three CSV files that mirror the Citrix inventory exports:

  Citrix command            Horizon equivalent            Output file
  ------------------------  ----------------------------  ------------------------
  Get-BrokerCatalog         Desktop pools (+ RDS pools)   HorizonDesktopPools.csv
  Get-ProvScheme            Pool/farm provisioning,       HorizonProvisioning.csv
                            golden image VM + snapshot
  Get-AcctIdentityPool      Pool/farm naming patterns     HorizonNamingPatterns.csv

Column notes for HorizonDesktopPools.csv (these differ from Citrix):
  AvailableCount   Machines whose Horizon state is AVAILABLE (idle and ready
                   for a user). This is NOT the Citrix catalog AvailableCount,
                   which counts machines not yet in a delivery group.
  UnassignedCount  Machines with no assigned user. Only filled in for DEDICATED
                   pools; blank for floating and RDS pools, where machines are
                   never assigned to a user.

CSV files are written as UTF-8 with a BOM so Excel shows non-ASCII names correctly.

Requirements: Python 3.8+, `pip install requests`
Permissions:  A Horizon admin account with read access to pools, farms,
              machines and vCenter config (the built-in "Administrators
              (Read only)" role is sufficient).

Usage:
  python omnissa_horizon_inventory.py --server cs01.corp.local \
      --username svc_horizon --domain CORP --output-dir C:\\HorizonExport

  --username also accepts DOMAIN\\user or user@domain.

Password is taken from --password (not recommended: visible in shell history
and the process list), the HZN_PASSWORD environment variable, or an
interactive prompt, in that order.
"""

import argparse
import csv
import getpass
import os
import sys
from collections import defaultdict

try:
    import requests
except ImportError:
    sys.exit("The 'requests' module is required:  pip install requests")


# Highest version first; the script falls back to lower versions automatically
# when a Connection Server doesn't support one (older Horizon builds).
POOL_VERSIONS = ["v8", "v7", "v6", "v5", "v4", "v3", "v2", "v1"]
FARM_VERSIONS = ["v5", "v4", "v3", "v2", "v1"]
MACHINE_VERSIONS = ["v4", "v3", "v2", "v1"]
RDS_VERSIONS = ["v2", "v1"]
VC_VERSIONS = ["v2", "v1"]

# Status codes that mean "this API version doesn't exist here" -> try the next one.
VERSION_MISSING_CODES = (400, 404, 405)

PAGE_SIZE = 1000
DEFAULT_TIMEOUT = 300

POOL_FIELDS = [
    "PoolName", "DisplayName", "PoolType", "ProvisioningType", "Source",
    "SessionSupport", "UserAssignment", "PersistUserChanges", "Enabled",
    "FarmName", "TotalMachines", "AvailableCount", "UsedCount",
    "UnassignedCount", "PoolId"]
PROV_FIELDS = [
    "Name", "ObjectType", "ProvisioningType", "Source", "FullClone",
    "vCenter", "GoldenImageVM", "GoldenImagePath", "Snapshot", "SnapshotPath",
    "VMTemplate", "NamingPattern", "ParentVmId", "BaseSnapshotId",
    "VmTemplateId", "Id"]
NAME_FIELDS = [
    "Name", "ObjectType", "NamingMethod", "NamingPattern", "MaxMachines",
    "MinMachines", "SpareMachines", "ProvisioningTime", "Id"]


def warn(msg):
    print(f"  WARNING: {msg}", file=sys.stderr)


def first(*values):
    """Return the first value that is not None or an empty string ('' if none)."""
    for v in values:
        if v is not None and v != "":
            return v
    return ""


class HorizonClient:
    def __init__(self, server, verify=True, timeout=DEFAULT_TIMEOUT):
        self.base = f"https://{server}/rest"
        self.session = requests.Session()
        # verify is passed per request: a session-level verify=False is silently
        # overridden when REQUESTS_CA_BUNDLE / CURL_CA_BUNDLE is set.
        self.verify = verify
        self.timeout = timeout
        self.refresh_token = None

    # ---- authentication -------------------------------------------------
    def login(self, username, password, domain):
        r = self.session.post(
            f"{self.base}/login",
            json={"username": username, "password": password, "domain": domain},
            timeout=self.timeout,
            verify=self.verify,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Login failed ({r.status_code}): {r.text}")
        data = r.json()
        self.refresh_token = data.get("refresh_token")
        self.session.headers["Authorization"] = f"Bearer {data['access_token']}"

    def _refresh(self):
        if not self.refresh_token:
            return False
        try:
            r = self.session.post(
                f"{self.base}/refresh",
                json={"refresh_token": self.refresh_token},
                timeout=self.timeout,
                verify=self.verify,
            )
        except requests.RequestException:
            return False
        if r.status_code != 200:
            return False
        self.session.headers["Authorization"] = f"Bearer {r.json()['access_token']}"
        return True

    def logout(self):
        if not self.refresh_token:
            return
        # Refresh first so logout still works if the access token expired during collection.
        self._refresh()
        try:
            self.session.post(
                f"{self.base}/logout",
                json={"refresh_token": self.refresh_token},
                timeout=self.timeout,
                verify=self.verify,
            )
        except requests.RequestException:
            pass

    # ---- HTTP helpers ---------------------------------------------------
    def get(self, path, params=None):
        r = self.session.get(f"{self.base}/{path}", params=params, timeout=self.timeout, verify=self.verify)
        if r.status_code == 401 and self._refresh():  # access token expired
            r = self.session.get(f"{self.base}/{path}", params=params, timeout=self.timeout, verify=self.verify)
        return r

    def get_paged(self, path, params=None):
        """GET a list endpoint, following page/size pagination."""
        results, page, last_first = [], 1, None
        while True:
            p = dict(params or {})
            p.update({"page": page, "size": PAGE_SIZE})
            r = self.get(path, p)
            r.raise_for_status()
            batch = r.json() or []
            # Guard against servers that ignore paging and return the same page
            if not batch or batch[0].get("id") == last_first:
                break
            results.extend(batch)
            if len(batch) < PAGE_SIZE:
                break
            last_first = batch[0].get("id")
            page += 1
        return results

    def get_versioned_list(self, area, resource, versions, paged=True, params=None):
        """Try each API version (newest first) until one is supported."""
        last_error = None
        for v in versions:
            path = f"{area}/{v}/{resource}"
            try:
                data = self.get_paged(path, params) if paged else self._get_json(path, params)
                return data, v
            except requests.HTTPError as e:
                # Only "version not supported" responses fall through to the next
                # version. Auth failures and 5xx are real errors and are raised.
                code = e.response.status_code if e.response is not None else None
                if code not in VERSION_MISSING_CODES:
                    raise
                last_error = e
        raise RuntimeError(f"No supported version for {area}/{resource}: {last_error}")

    def _get_json(self, path, params=None):
        r = self.get(path, params)
        r.raise_for_status()
        return r.json()


class VCenterLookup:
    """Resolves vCenter object IDs (vm-123, snapshot-4) to names/paths, with caching."""

    def __init__(self, client):
        self.c = client
        self.vc_names = {}
        self.base_vms = {}     # vcenter_id -> {vm_id: info}
        self.templates = {}    # vcenter_id -> {vm_id: info}
        self.snapshots = {}    # (vcenter_id, vm_id) -> {snap_id: info}
        self.warned = set()

    def load_vcenters(self):
        try:
            vcs, _ = self.c.get_versioned_list("config", "virtual-centers", VC_VERSIONS, paged=False)
            for vc in vcs:
                self.vc_names[vc.get("id")] = first(vc.get("server_name"), vc.get("name"), vc.get("id"))
        except Exception as e:  # noqa: BLE001
            warn(f"could not list vCenters ({e}); vCenter IDs will be shown instead")

    def _warn_once(self, key, msg):
        if key not in self.warned:
            self.warned.add(key)
            warn(msg)

    def vcenter(self, vc_id):
        return self.vc_names.get(vc_id, vc_id or "")

    def base_vm(self, vc_id, vm_id):
        if not vc_id or not vm_id:
            return {}
        if vc_id not in self.base_vms:
            try:
                vms = self.c._get_json("external/v1/base-vms", {"vcenter_id": vc_id})
                self.base_vms[vc_id] = {v["id"]: v for v in vms}
            except Exception as e:  # noqa: BLE001
                self._warn_once(("vm", vc_id), f"base VM lookup failed for vCenter {vc_id}: {e}")
                self.base_vms[vc_id] = {}
        return self.base_vms[vc_id].get(vm_id, {})

    def template(self, vc_id, vm_id):
        if not vc_id or not vm_id:
            return {}
        if vc_id not in self.templates:
            try:
                ts = self.c._get_json("external/v1/vm-templates", {"vcenter_id": vc_id})
                self.templates[vc_id] = {t["id"]: t for t in ts}
            except Exception as e:  # noqa: BLE001
                self._warn_once(("tpl", vc_id), f"template lookup failed for vCenter {vc_id}: {e}")
                self.templates[vc_id] = {}
        return self.templates[vc_id].get(vm_id, {})

    def snapshot(self, vc_id, vm_id, snap_id):
        if not vc_id or not vm_id or not snap_id:
            return {}
        key = (vc_id, vm_id)
        if key not in self.snapshots:
            try:
                snaps = self.c._get_json(
                    "external/v1/base-snapshots", {"vcenter_id": vc_id, "base_vm_id": vm_id}
                )
                self.snapshots[key] = {s["id"]: s for s in snaps}
            except Exception as e:  # noqa: BLE001
                self._warn_once(("snap",) + key, f"snapshot lookup failed for {vm_id}: {e}")
                self.snapshots[key] = {}
        return self.snapshots[key].get(snap_id, {})


# ---- mapping helpers ------------------------------------------------------

def provisioning_type(source, pool_type):
    """Translate Horizon 'source' into a Citrix-style provisioning label."""
    s = (source or "").upper()
    if pool_type == "RDS":
        return "RDS Farm"
    if "INSTANT_CLONE" in s:
        return "Instant Clone"
    if s == "VIEW_COMPOSER":
        return "Linked Clone (Composer)"
    if s == "VIRTUAL_CENTER":
        return "Full Clone" if pool_type == "AUTOMATED" else "Manual (vCenter VMs)"
    if s == "UNMANAGED":
        return "Manual (Physical/Unmanaged)"
    return s or pool_type or ""


def machine_users(m):
    return m.get("user_ids") or m.get("assigned_user_ids") or []


USED_STATES = {
    "CONNECTED", "DISCONNECTED", "UNASSIGNED_USER_CONNECTED",
    "UNASSIGNED_USER_DISCONNECTED", "ALREADY_USED", "IN_PROGRESS",
}


def write_csv(path, rows, fields):
    # utf-8-sig adds a BOM so Excel detects UTF-8 and shows non-ASCII names correctly.
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    print(f"  Wrote {len(rows):>4} rows -> {path}")


# ---- main collection ------------------------------------------------------

def collect(client, out_dir):
    print("Collecting desktop pools...")
    pools, pv = client.get_versioned_list("inventory", "desktop-pools", POOL_VERSIONS)
    print(f"  {len(pools)} pools (API {pv})")

    print("Collecting farms...")
    try:
        farms, fv = client.get_versioned_list("inventory", "farms", FARM_VERSIONS)
        print(f"  {len(farms)} farms (API {fv})")
    except Exception as e:  # noqa: BLE001
        warn(f"could not list farms ({e})")
        farms = []
    farm_by_id = {f.get("id"): f for f in farms}

    print("Collecting machines (for pool counts)...")
    machines, mv = client.get_versioned_list("inventory", "machines", MACHINE_VERSIONS)
    print(f"  {len(machines)} machines (API {mv})")
    by_pool = defaultdict(list)
    for m in machines:
        by_pool[m.get("desktop_pool_id")].append(m)

    rds_by_farm = defaultdict(list)
    if farms:
        try:
            rds, _ = client.get_versioned_list("inventory", "rds-servers", RDS_VERSIONS)
            for s in rds:
                rds_by_farm[s.get("farm_id")].append(s)
        except Exception as e:  # noqa: BLE001
            warn(f"could not list RDS servers ({e})")

    vc = VCenterLookup(client)
    vc.load_vcenters()

    pool_rows, prov_rows, name_rows = [], [], []

    # ---- desktop pools ----
    for p in pools:
        ptype = p.get("type", "")
        source = p.get("source", "")
        farm = farm_by_id.get(p.get("farm_id"), {})
        ms = by_pool.get(p.get("id"), [])
        assignment = p.get("user_assignment", "")

        if ptype == "RDS":
            servers = rds_by_farm.get(p.get("farm_id"), [])
            total, available, used, unassigned = len(servers), "", "", ""
        else:
            total = len(ms)
            available = sum(1 for m in ms if m.get("state") == "AVAILABLE")
            used = sum(1 for m in ms if m.get("state") in USED_STATES)
            # Floating machines are never assigned to a user, so the count is
            # only meaningful for DEDICATED pools.
            unassigned = sum(1 for m in ms if not machine_users(m)) \
                if assignment == "DEDICATED" else ""

        pool_rows.append({
            "PoolName": p.get("name", ""),
            "DisplayName": p.get("display_name", ""),
            "PoolType": ptype,
            "ProvisioningType": provisioning_type(source, ptype),
            "Source": source,
            "SessionSupport": "MultiSession" if ptype == "RDS" else "SingleSession",
            "UserAssignment": assignment or ("SHARED (RDS)" if ptype == "RDS" else ""),
            "PersistUserChanges": "Persistent (dedicated)" if assignment == "DEDICATED"
                                  else "Non-persistent" if assignment == "FLOATING" else "",
            "Enabled": p.get("enabled", ""),
            "FarmName": farm.get("name", ""),
            "TotalMachines": total,
            "AvailableCount": available,
            "UsedCount": used,
            "UnassignedCount": unassigned,
            "PoolId": p.get("id", ""),
        })

        if ptype == "AUTOMATED":
            prov_rows.append(provisioning_row("DesktopPool", p, source,
                                              p.get("provisioning_settings") or {},
                                              p.get("pattern_naming_settings") or {},
                                              p.get("vcenter_id"), vc))
            name_rows.append(naming_row("DesktopPool", p,
                                        p.get("naming_method"),
                                        p.get("pattern_naming_settings") or {},
                                        p.get("specific_naming_settings")
                                        or p.get("specified_naming_settings") or {}))

    # ---- automated RDS farms ----
    for f in farms:
        if f.get("type") != "AUTOMATED":
            continue
        auto = f.get("automated_farm_settings") or {}
        prov_rows.append(provisioning_row("Farm", f, f.get("source", ""),
                                          auto.get("provisioning_settings") or {},
                                          auto.get("pattern_naming_settings") or {},
                                          auto.get("vcenter_id") or f.get("vcenter_id"), vc))
        name_rows.append(naming_row("Farm", f, "PATTERN",
                                    auto.get("pattern_naming_settings") or {}, {}))

    os.makedirs(out_dir, exist_ok=True)
    print("Writing CSV files...")
    write_csv(os.path.join(out_dir, "HorizonDesktopPools.csv"), pool_rows, POOL_FIELDS)
    write_csv(os.path.join(out_dir, "HorizonProvisioning.csv"), prov_rows, PROV_FIELDS)
    write_csv(os.path.join(out_dir, "HorizonNamingPatterns.csv"), name_rows, NAME_FIELDS)


def provisioning_row(obj_type, obj, source, prov, naming, vc_id, vc):
    parent_id = prov.get("parent_vm_id")
    snap_id = prov.get("base_snapshot_id")
    tpl_id = prov.get("vm_template_id")
    vm = vc.base_vm(vc_id, parent_id)
    snap = vc.snapshot(vc_id, parent_id, snap_id)
    tpl = vc.template(vc_id, tpl_id)
    return {
        "Name": obj.get("name", ""),
        "ObjectType": obj_type,
        "ProvisioningType": provisioning_type(source, "AUTOMATED"),
        "Source": source,
        "FullClone": (source or "").upper() == "VIRTUAL_CENTER",
        "vCenter": vc.vcenter(vc_id),
        "GoldenImageVM": first(vm.get("name"), parent_id),
        "GoldenImagePath": first(vm.get("path")),
        "Snapshot": first(snap.get("name"), snap_id),
        "SnapshotPath": first(snap.get("path")),
        "VMTemplate": first(tpl.get("name"), tpl_id),
        "NamingPattern": first(naming.get("naming_pattern")),
        "ParentVmId": first(parent_id),
        "BaseSnapshotId": first(snap_id),
        "VmTemplateId": first(tpl_id),
        "Id": obj.get("id", ""),
    }


def naming_row(obj_type, obj, method, pattern, specified):
    if not method:
        method = "PATTERN" if pattern else ("SPECIFIED" if specified else "")
    return {
        "Name": obj.get("name", ""),
        "ObjectType": obj_type,
        "NamingMethod": method,
        "NamingPattern": first(pattern.get("naming_pattern")),
        "MaxMachines": first(pattern.get("max_number_of_machines"),
                             pattern.get("max_number_of_rds_servers")),
        "MinMachines": first(pattern.get("min_number_of_machines")),
        "SpareMachines": first(pattern.get("number_of_spare_machines"),
                               specified.get("number_of_spare_machines")),
        "ProvisioningTime": first(pattern.get("provisioning_time")),
        "Id": obj.get("id", ""),
    }


def split_user(username, domain):
    """Accept user, DOMAIN\\user or user@domain; return (user, domain)."""
    if "\\" in username:
        d, u = username.split("\\", 1)
        if domain and domain.lower() != d.lower():
            warn(f"using domain '{d}' from the user name; --domain '{domain}' is ignored")
        return u, d
    if "@" in username:
        u, d = username.split("@", 1)
        if not domain:
            domain = d.split(".")[0]
            warn(f"domain guessed as '{domain}' from the UPN; if login fails, "
                 "pass --domain with the NetBIOS name")
        return u, domain
    return username, domain


def main():
    ap = argparse.ArgumentParser(description="Export Omnissa Horizon inventory to CSV")
    ap.add_argument("--server", required=True, help="Connection Server FQDN")
    ap.add_argument("--username", required=True, help="user, DOMAIN\\user or user@domain")
    ap.add_argument("--domain", help="AD domain (NetBIOS name); optional if given in --username")
    ap.add_argument("--password", help="Not recommended; prefer HZN_PASSWORD or the prompt")
    ap.add_argument("--output-dir", default=".", help="Folder for the CSV files")
    ap.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT,
                    help=f"Per-request timeout in seconds (default {DEFAULT_TIMEOUT})")
    ap.add_argument("--ca-bundle", help="Path to a CA bundle to trust (PEM)")
    ap.add_argument("--insecure", action="store_true",
                    help="Skip TLS certificate validation (lab/self-signed only)")
    args = ap.parse_args()

    username, domain = split_user(args.username, args.domain)
    if not domain:
        ap.error("specify --domain or enter the user as DOMAIN\\user")

    if args.password:
        warn("--password is visible in shell history and the process list; prefer HZN_PASSWORD")
    password = args.password or os.environ.get("HZN_PASSWORD") or \
        getpass.getpass(f"Password for {domain}\\{username}: ")

    verify = False if args.insecure else (args.ca_bundle or True)
    if args.insecure:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    client = HorizonClient(args.server, verify=verify, timeout=args.timeout)
    try:
        print(f"Logging in to {args.server} as {domain}\\{username}...")
        client.login(username, password, domain)
        try:
            collect(client, args.output_dir)
        finally:
            client.logout()
    except (RuntimeError, requests.RequestException, OSError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    print("Done.")


if __name__ == "__main__":
    main()
