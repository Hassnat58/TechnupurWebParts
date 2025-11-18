import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import './style.scss';

interface INavigationItem {
  id: number;
  title: string;
  url: string;
  parent: string;
  order: number;
  isActive: boolean;
  description: string;
  children: INavigationItem[];
}

export default class HeaderExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<any> {

  public topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context as any
    });
    this.context.placeholderProvider.changedEvent.add(this, this.renderCustomHeader);

    // Call render method for generating the HTML elements.  
    this.renderCustomHeader();

    return Promise.resolve();
  }

  public async fetchNavigationData(): Promise<INavigationItem[]> {
    try {
      const items = await sp.web.lists.getByTitle("NavigationMenu")
        .items.select(
          "ID",
          "GP_PrimaryLinkTitle",
          "GP_PrimaryLinkURL",
          "GP_PrimaryParent",
          "GP_Order",
          "GP_IsActive",
          "GP_Description"
        )
        .filter("GP_IsActive eq 'Yes'")
        .orderBy("GP_Order", true)
        .get();

      return this.buildNavigationHierarchy(items);
    } catch (error) {
      console.error("Error fetching navigation:", error);
      return [];
    }
  }

  public buildNavigationHierarchy(items: any[]): INavigationItem[] {
    const itemsMap = new Map<string, INavigationItem>();
    const rootItems: INavigationItem[] = [];

    items.forEach(item => {
      itemsMap.set(item.GP_PrimaryLinkTitle, {
        id: item.ID,
        title: item.GP_PrimaryLinkTitle,
        url: item.GP_PrimaryLinkURL?.Url || '',
        parent: item.GP_PrimaryParent,
        order: item.GP_Order || 0,
        isActive: item.GP_IsActive === 'Yes',
        description: item.GP_Description || '',
        children: []
      });
    });

    items.forEach(item => {
      const currentItem = itemsMap.get(item.GP_PrimaryLinkTitle);
      if (currentItem) {
        if (item.GP_PrimaryParent) {
          const parentItem = itemsMap.get(item.GP_PrimaryParent);
          if (parentItem) {
            parentItem.children.push(currentItem);
          }
        } else {
          rootItems.push(currentItem);
        }
      }
    });

    return this.sortNavigationItems(rootItems);
  }

  public sortNavigationItems(items: INavigationItem[]): INavigationItem[] {
    items.sort((a, b) => a.order - b.order);
    items.forEach(item => {
      if (item.children.length > 0) {
        item.children = this.sortNavigationItems(item.children);
      }
    });
    return items;
  }

  public generateNavigationHTML(items: INavigationItem[]): string {
    return items.map(item => {
      const hasChildren = item.children.length > 0;
      const mainLinkOnClick = item.url ? ` onclick="window.open('${item.url}', '_blank'); return false;"` : '';
      const titleWithDescription = item.description ? ` title="${item.description}"` : '';

      return `
        <li class="${hasChildren ? 'has-dropdown' : ''}"${titleWithDescription}>
          <a href="javascript:void(0);"${mainLinkOnClick}>
            ${item.title}${hasChildren ? ' <span class="dropdown-arrow">▼</span>' : ''}
          </a>
          ${hasChildren ? `
            <ul class="dropdown">
              ${this.generateNavigationHTML(item.children)}
            </ul>
          ` : ''}
        </li>
      `;
    }).join('');
  }

  public async renderCustomHeader(): Promise<void> {
    if (this.topPlaceholder) return;

    this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    if (!this.topPlaceholder?.domElement) return;

    const imageUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/GPLogo.png`;
    const navigationItems = await this.fetchNavigationData();
    const currentUser = await this.fetchCurrentUser();

    this.topPlaceholder.domElement.innerHTML = `
      <header class="TopHead">
        <!-- your logo -->
        <div class="MyLogo">
          <img src="${imageUrl}" alt="Logo">
        </div>
        
        <!-- menu -->
        <div class="MyNav navbar">
          ${navigationItems.map(item => {
            if (item.children.length === 0) {
              return `<a href="${item.url || '#'}">${item.title}</a>`;
            } else {
              return `
                <div class="dropdown">
                  <button class="dropbtn">${item.title}<i class="Arrow"><img src="https://static.thenounproject.com/png/2424963-512.png" /></i></button>
                  <div class="dropdown-content">
                    <div class="header">${item.title}</div>
                    ${this.generateMegaMenuColumns(item.children)}
                  </div>
                </div>
              `;
            }
          }).join('')}
        </div>
        
        <!-- user div -->
        <div class="MyUser">
          ${currentUser ? `
            <div class="user-section">
              <div class="user-info">
                <span class="user-name">${currentUser.name}</span>
                <img src="${currentUser.picture}" alt="${currentUser.name}" class="user-image">
              </div>
            </div>
          ` : ''}
        </div>
        
        <div style="clear:both;">
      </header>
      
      ${this.getCustomStyles()}
    `;
  }

  private getCustomStyles(): string {
    return `
      <style>
        * {
          box-sizing: border-box;
        }

        body {
          margin: 0;
        }

      .MyNav {
        font-family: 'Barlow'; /* Apply font-family to the main navigation */
      }

        .MyNav.navbar {
          overflow: hidden;
          font-family: 'Barlow';
        }

        .MyNav.navbar ul,
        .MyNav.navbar ul li {
          margin:0;
          padding:0;
          list-style:none;
        }

        .MyNav.navbar a {
          /*float: left;*/
          display:inline-block;
          font-size: 14px;
          color: #000;
          text-align: center;
          padding: 14px 16px;
          text-decoration: none;
        }

        .MyNav .dropdown {
          /*float: left;*/
          display:inline;
          overflow: hidden;
        }

        .MyNav .dropdown .dropbtn {
          font-size: 16px;  
          border: none;
          outline: none;
          color: #000;
          padding: 14px 16px;
          background-color: inherit;
          font: inherit;
          margin: 0;
          font-family: 'Barlow';
        }

        .user-section {
              position: absolute;
              right: 20px;
              z-index: 2;
            }
        .user-info {
              display: flex;
              align-items: center;
              gap: 10px;
            }

        .user-name {
              color: #2A3440;
              font-family: 'Barlow';
              font-size: 14px;
              font-weight: 500;
            }

        .user-image {
              width: 32px;
              height: 32px;
              border-radius: 50%;
              object-fit: cover;
            }

        .MyNav .dropdown .dropbtn img {
          height:12px;
          position:relative;
          top:2px;
          margin-left:8px;
        }

        .MyNav.navbar a:hover,
        .MyNav .dropdown:hover .dropbtn {
          color: #0078d4;
          /*background-color: red;*/
          font-family: 'Barlow';
        }

        /* DropDown WhiteBox to Show Items */
        .MyNav .dropdown-content {
          display: none;
          position: absolute;
          background-color: #fff;
          width: 100%;
          left: 0;
          box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.1);
          z-index: 1;
          font-family: 'Barlow';
        }

        .MyNav .dropdown-content .header {
          padding: 20px;
          color: #000;
          width:100%;
          display:block;
          clear:both;
          border-top:1px solid #ccc;
          background:#f2f2f2;
          text-align:left;
          font-family: 'Barlow';
        }

        .MyNav .dropdown:hover .dropdown-content {
          display: block;
          text-align:left;
        }

        /* Create three equal columns that floats next to each other */
        .MyNav .column {
          float: left;
          width: 33.33%;
          padding: 10px;
        }

        .MyNav .column a {
          float: none;
          color: black;
          padding: 16px;
          text-decoration: none;
          display: block;
          text-align: left;
          font-weight:bold;
          font-size:16px;
        }

        .MyNav .column a:hover {
          color: #0078d4;
        }

        /* Clear floats after the columns */
        .MyNav .row:after {
          content: "";
          display: table;
          clear: both;
          
        }


        .TopHead {
          display:block;
          width:100%;
          border-bottom:1px solid #ccc;
          padding:10px 20px;
          background-color: white;
          color: #2A3440;
          height: 57px;
          align-items: center;
        }
        .TopHead .MyLogo {
          display:inline-block;
          float:left;
          width:20%;
        }
        .TopHead .MyLogo img { height:45px; width:93.19px; }

        .TopHead .MyNav {
          display:inline-block;
          width:60%;
          text-align:center;
        }

        .TopHead .MyUser {
          display:inline-block;
          float:right;
          width:20%;
          text-align:right;
          padding-top: 7px;
        }

        .MyNav {
          font-family: 'Barlow';
        }

        .MyNav .column .column-header {
          font-weight: bold;
          font-size: 16px;
          margin-bottom: 10px;
          display: block;
          color: #000;
          font-family: 'Barlow';
        }

        .MyNav .column .sub-items a {
          font-weight: normal;
          color: #000;
          text-decoration: none;
          font-family: 'Barlow';
        }

        .MyNav .column .sub-items a:hover {
          color: red;
          font-family: 'Barlow';
        }
      </style>
    `;
  }

  private generateMegaMenuColumns(children: INavigationItem[]): string {
    return `
      <div class="row">
        ${children.map(child => `
          <div class="column">
            <a href="${child.url || '#'}" class="column-header">${child.title}</a>
            ${child.children && child.children.length > 0 ? `
              <ul class="sub-items">
                ${child.children.map(subItem => `
                  <li><a href="${subItem.url || '#'}" title="${subItem.description || ''}" class="sub-item">${subItem.title}</a></li>
                `).join('')}
              </ul>
            ` : ''}
          </div>
        `).join('')}
      </div>
    `;
  }

  public async fetchCurrentUser() {
    try {
      const user = await sp.web.currentUser.get();
      return {
        name: user.Title,
        email: user.Email,
        picture: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${user.Email}`
      };
    } catch (error) {
      console.error("Error fetching user:", error);
      return null;
    }
  }
}
