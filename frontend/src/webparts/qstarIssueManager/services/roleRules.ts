import { Role } from "../models/IRole";

export interface IRoleGroupNames {
  admin: string;
  qm: string;
  owner: string;
  reader: string;
}

export const DEFAULT_ROLE_GROUP_NAMES: IRoleGroupNames = {
  admin: "Q-Star Admins",
  qm: "Q-Star Quality Managers",
  owner: "Q-Star Task Owners",
  reader: "Q-Star Readers",
};

export interface IRoleMatch {
  role: Role;
  matchedGroup?: string;
}

export interface ISitePermissionSnapshot {
  manageWeb: boolean;
  editListItems: boolean;
}

export function resolveRoleFromGroupTitles(
  groupTitles: string[],
  groupNames: IRoleGroupNames = DEFAULT_ROLE_GROUP_NAMES
): IRoleMatch {
  const titles = new Set(groupTitles.map((title) => title.toLocaleLowerCase()));
  const priority: Role[] = ["admin", "qm", "owner", "reader"];

  for (const role of priority) {
    const expected = groupNames[role];
    if (titles.has(expected.toLocaleLowerCase())) {
      return { role, matchedGroup: expected };
    }
  }

  return { role: "reader" };
}

export function resolveBetaRoleFromSitePermissions(permissions: ISitePermissionSnapshot): IRoleMatch {
  if (permissions.manageWeb) {
    return { role: "admin", matchedGroup: "Site owner permissions (beta)" };
  }
  if (permissions.editListItems) {
    return { role: "qm", matchedGroup: "Site edit permissions (beta)" };
  }
  return { role: "reader", matchedGroup: "Site read permissions (beta)" };
}
