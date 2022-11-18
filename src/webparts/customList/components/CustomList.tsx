import * as React from "react";
import styles from "./CustomList.module.scss";
import { ICustomListProps } from "./ICustomListProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { useState, useEffect } from "react";
import CustomService from "../services/CustomService";

import { sp } from "@pnp/sp/presets/all";

export default class CustomList extends React.Component<ICustomListProps, {}> {
  private customService: CustomService;

  public constructor(props: ICustomListProps, state: {}) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any,
    });

    this.customService = new CustomService();
    this.customService.webUrl = this.props.siteUrl;
    this.customService.listName = "ExcelLoop";
    this.state = {
      data: {},
    };
    this.getList();
  }

  async getList() {
    var model = await this.customService.getlist(
      this.props.context,
      2,
      100,
      "ID,Title",
      (res: any) => {
        debugger;
      }
    );
  }

  public render(): React.ReactElement<ICustomListProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      siteUrl,
    } = this.props;

    return (
      <section
        className={`${styles.customList} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint lndgjhajhgkjlasg Framework (SPFx) is a extensibility
            model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s
            the easiest way to extend Microsoft 365 with automatic Single Sign
            On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li>
              <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
                SharePoint Framework Overview
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-graph"
                target="_blank"
                rel="noreferrer"
              >
                Use Microsoft Graph in your solution
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-teams"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Teams using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-viva"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Viva Connections using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-store"
                target="_blank"
                rel="noreferrer"
              >
                Publish SharePoint Framework applications to the marketplace
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-api"
                target="_blank"
                rel="noreferrer"
              >
                SharePoint Framework API reference
              </a>
            </li>
            <li>
              <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
                Microsoft 365 Developer Community
              </a>
            </li>
          </ul>
        </div>
      </section>
    );
  }
}