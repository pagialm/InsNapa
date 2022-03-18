import { IGroup } from "office-ui-fabric-react";

export interface IDashboardState {
  items: IListItem[];
  allItems?: IListItem[];
  groups?:IGroup[];
  allGroups?: IGroup[];
}
