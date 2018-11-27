export default class AnalyticsConfig {
  public static async getTrackingId(): Promise<string> {
    return "UA-129865562-1";
  }

  public async setup(): Promise<boolean> {
    return true;
  }

  private async ensureConfig(): Promise<boolean> {
    return true;
  }
}
