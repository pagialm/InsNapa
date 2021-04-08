import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IInsuranceNapaProps {
  description: string;
  context: WebPartContext;
  itemId?: number;
}
