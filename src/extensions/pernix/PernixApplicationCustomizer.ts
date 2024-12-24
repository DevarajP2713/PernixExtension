import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { sp } from "@pnp/sp/presets/all";
import "./style.css";

export interface IPernixApplicationCustomizerProperties {
  testMessage: string;
}

export default class PernixApplicationCustomizer extends BaseApplicationCustomizer<IPernixApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context as unknown as undefined,
      });

      this._getAdmins();
    });
  }

  async _getAdmins() {
    const response: any = await sp.web.currentUser.get();
    const curUserEmail: string = response?.Email?.toLowerCase() ?? "";

    const superAdminUsers: any[] = await sp.web.siteGroups
      .getByName("Pernix_Connect_Super_Admin")
      .users.get();

    // const headerAdminUsers: any[] = await sp.web.siteGroups
    //   .getByName("Header_Admin")
    //   .users.get();

    const arrayAdmins: any[] = [...superAdminUsers];
    // const arrayAdmins: any[] = [...superAdminUsers, ...headerAdminUsers];
    const isAdmin: boolean = arrayAdmins?.some(
      (val: any) => val?.Email.toLowerCase() === curUserEmail
    );

    const siteHeader = document.querySelector(".sp-pageLayout-horizontalNav");
    const siteContent = document.querySelector(".sp-App-bodyContainer");

    if (siteHeader) {
      if (!isAdmin) {
        siteHeader.setAttribute("data-custom-class", "nonAdmin");
      } else {
        siteHeader.removeAttribute("data-custom-class");
      }
    }

    if (siteContent) {
      if (!isAdmin) {
        siteContent.setAttribute("data-custom-class", "nonAdmin");
      } else {
        siteContent.removeAttribute("data-custom-class");
      }
    }

    return Promise.resolve();
  }
}
