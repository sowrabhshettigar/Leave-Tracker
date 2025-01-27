import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import LeaveTracker from "./components/LeaveTracker";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface ILeaveTrackerWebPartProps {
  itemId?: string; // Add this property to support edit mode
}

export default class LeaveTrackerWebPart extends BaseClientSideWebPart<ILeaveTrackerWebPartProps> {
  public render(): void {
    const element: React.ReactElement = React.createElement(LeaveTracker, {
      context: this.context,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      itemId: this.properties.itemId ? parseInt(this.properties.itemId) : undefined
    });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
