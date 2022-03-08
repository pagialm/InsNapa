import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFieldSP } from "./IFieldSP";

export interface IFilteredDropdownProps {
  //   options: IDropdownOption[];
  listname: string;
  field1: IFieldSP;
  field2?: IFieldSP;
  field3?: IFieldSP;
  filterValue?: string;
  label: string;
  context: WebPartContext;
}
