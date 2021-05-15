import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';


export interface IHelloWorldWebPartProps {
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _uniqueId;
  private exposePageContextGlobally = true;
  private exposeTeamsContextGlobally = true;

  public render(): void {
    this._uniqueId = this.context.instanceId;
    this.domElement.innerHTML = `$HTML$`;
    this.executeScript(this.domElement);
  }

  // src: https://github.com/pnp/sp-dev-fx-webparts/blob/4eede437dcdefa8a4416698529119760abf57643/samples/react-script-editor/src/webparts/scriptEditor/ScriptEditorWebPart.ts#L52
  private async executeScript(element: HTMLElement) {
    // clean up added script tags in case of smart re-load        
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    let scriptTags = headTag.getElementsByTagName("script");
    for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this._uniqueId) {
        headTag.removeChild(scriptTag);
      }
    }

    if (this.exposePageContextGlobally && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    if (this.exposeTeamsContextGlobally && !window["_teamsContextInfo"]) {
      window["_teamsContextInfo"] = this.context.sdks.microsoftTeams.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    // main section of function
    const scripts = [];
    const children_nodes = element.getElementsByTagName("script");

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        // Add unique param to force load on each run to overcome smart navigation in the browser as needed
        const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
        scriptUrl += prefix + 'pnp=' + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        if (console.error) {
          console.error(error);
        }
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      // Copies all attributes in case of loaded script relies on the tag attributes
      if (attr.name.toLowerCase() === "onload") continue; // onload handled after loading with SPComponentLoader
      scriptTag.setAttribute(attr.name, attr.value);
    }

    // set a bogus type to avoid browser loading the script, as it's loaded with SPComponentLoader
    scriptTag.type = (scriptTag.src && scriptTag.src.length) > 0 ? "pnp" : "text/javascript";
    // Ensure proper setting and adding id used in cleanup on reload
    scriptTag.setAttribute("pnpname", this._uniqueId);

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
