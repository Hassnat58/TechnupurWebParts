import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IFeelingtaskProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  spInstance: SPFI;
}