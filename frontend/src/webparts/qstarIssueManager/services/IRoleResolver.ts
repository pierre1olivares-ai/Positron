import { IResolvedRole } from "../models/IRole";

export interface IRoleResolver {
  resolve(): Promise<IResolvedRole>;
}
