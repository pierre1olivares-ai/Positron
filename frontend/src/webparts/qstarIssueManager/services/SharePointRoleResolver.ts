import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/security/web";
import { PermissionKind } from "@pnp/sp/security";

import { IResolvedRole, Role } from "../models/IRole";
import { IRoleResolver } from "./IRoleResolver";
import { DEFAULT_ROLE_GROUP_NAMES, IRoleGroupNames, resolveBetaRoleFromSitePermissions, resolveRoleFromGroupTitles } from "./roleRules";

type GroupInfo = { Title: string };

export class SharePointRoleResolver implements IRoleResolver {
  private sp: SPFI;

  constructor(
    context: WebPartContext,
    private groupNames: IRoleGroupNames = DEFAULT_ROLE_GROUP_NAMES,
    siteUrl?: string,
    private betaAccessMode: boolean = false
  ) {
    this.sp = siteUrl ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context));
  }

  public async resolve(): Promise<IResolvedRole> {
    const groups = await this.sp.web.currentUser.groups.select("Title")() as GroupInfo[];
    const match = resolveRoleFromGroupTitles(groups.map((group) => group.Title), this.groupNames);
    if (!match.matchedGroup && this.betaAccessMode) {
      const [manageWeb, editListItems] = await Promise.all([
        this.sp.web.currentUserHasPermissions(PermissionKind.ManageWeb),
        this.sp.web.currentUserHasPermissions(PermissionKind.EditListItems),
      ]);
      return { ...resolveBetaRoleFromSitePermissions({ manageWeb, editListItems }), source: "sharepoint" };
    }
    return { ...match, source: "sharepoint" };
  }
}

export class DevelopmentRoleResolver implements IRoleResolver {
  constructor(private role: Role = "admin") {}

  public async resolve(): Promise<IResolvedRole> {
    return { role: this.role, source: "development", matchedGroup: "Local development override" };
  }
}
