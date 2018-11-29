import { sp, Item } from "@pnp/sp";
import { dateAdd } from "@pnp/common";
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

export default class AnalyticsConfig {
  private readonly googleTrackingIdConfigKey: string = "google-tracking-id";
  private readonly configListName: string = "CustomSettings";
  private context: ApplicationCustomizerContext = null;

  public async setup(context: ApplicationCustomizerContext): Promise<void> {
    this.context = context;
    sp.setup({ spfxContext: this.context });
    await this.ensureConfigList();
  }

  public async getGoogleTrackingId(): Promise<string> {
    const configItem = await this.getConfigItem(this.googleTrackingIdConfigKey);
    return this.getConfigItemValue(configItem);
    // return "UA-129865562-1";
  }

  public async setGoogleTrackingId(trackingId: string): Promise<string> {
    let configItem = await this.getConfigItem(this.googleTrackingIdConfigKey);

    if (configItem) {
      configItem = await this.updateConfigItem(configItem, trackingId);
    }
    else {
      configItem = await this.addConfigItem(this.googleTrackingIdConfigKey, trackingId);
    }

    return this.getConfigItemValue(configItem);
  }

  private getConfigItemValue(configItem: any) {
    return configItem ? configItem.Value : "";
  }

  private async getConfigItem(key: string): Promise<Item> {
    const configList = sp.web.lists.getByTitle(this.configListName);

    //Leverage caching to reduce calls to config list
    const getConfigItemResult = await configList.items.filter(`Title eq '${key}'`).top(1)
      .usingCaching({
        key: `${this.configListName}|${key}|${this.context.pageContext.web.id}`,
        expiration: dateAdd(new Date(), "hour", 8), //8 hours session cache
        storeName: "session"
      }).get();

    return getConfigItemResult.length === 1 ? getConfigItemResult[0] : null;
  }

  private async addConfigItem(key: string, value: string): Promise<Item> {
    const configList = sp.web.lists.getByTitle(this.configListName);
    let addConfigItemResult = await configList.items.add({ Title: key, Value: value });
    return addConfigItemResult.item;
  }

  private async updateConfigItem(configItem: Item, value: string): Promise<Item> {
    let updateConfigItemResult = await configItem.update({ Value: value });
    return updateConfigItemResult.item;
  }

  private async ensureConfigList(): Promise<boolean> {
    try {
      try {
        await sp.web.lists.getByTitle(this.configListName).get();
      }
      catch (err) {
        let result = await sp.web.lists.add(this.configListName);
        await result.list.fields.getByTitle("Title").update({ Title: "Key" });
        await result.list.fields.addText("Value");
        await result.list.defaultView.fields.add("Value");
      }
      //Config list was either found or created successfully
      return true;
    }
    catch (err) {
      //Config list could not be found or wasn't able to be created successfully
      return false;
    }
  }
}
