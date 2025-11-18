import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOrgchartProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  employeeCount: number;
  _isShowingAll?: boolean;
}
