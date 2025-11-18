import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEventswebpartspfxProps {
  description: string;
  context: WebPartContext;
  siteurl: string;
  itemsToShow: number;
}
