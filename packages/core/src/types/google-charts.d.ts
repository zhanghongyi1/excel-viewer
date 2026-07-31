/**
 * google-charts npm 包类型声明
 *
 * ES6 封装的 Google Charts 加载器
 * @see https://www.npmjs.com/package/google-charts
 */
declare module 'google-charts' {
  /** Google Charts 管理器（单例） */
  export class GoogleCharts {
    /** Google API 对象 (window.google)，加载完成后可用 */
    static api: any;

    /**
     * 加载 Google Charts 库并执行回调
     * @param callback 加载完成后的回调函数
     * @param type 可选，指定额外加载的包名或配置
     */
    static load(callback: () => void, type?: string | string[] | object): Promise<void>;
  }
}
