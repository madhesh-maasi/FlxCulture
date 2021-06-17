import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FlxCultureWebPart.module.scss';
import * as strings from 'FlxCultureWebPartStrings';

import { sp } from "@pnp/sp/presets/all";
import "../../ExternalRef/css/style.css";
import "../../ExternalRef/css/bootstrap.min.css";
import * as $ from "jquery";

export interface IFlxCultureWebPartProps {
  description: string;
}

export default class FlxCultureWebPart extends BaseClientSideWebPart<IFlxCultureWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }    
  public render(): void {      
    this.domElement.innerHTML = `
    <div class="loader-section" style="display:none"> 
    <div class="loader"></div>  
    </div>  
    <div class="container container-sm container-lg contaoiner-md mx-4 bg-color">  
    <div class="tile-head bg-secondary p-2">
    <h6 class="mx-2 mt-2">FLX Culture</h6>
    </div>
    <div class="flexculture">
    <ul class="list-unstyled" id="tile">
    <li class="tile p-2">   
    <h6 class="mx-2 ">Carring</h6>          
    <p class="mx-2 ">Help one another when it comes complting the job</p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Communicate</h6>     
    <p class="mx-2 ">Open door and transparent</p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Diversity</h6>
    <p class="mx-2 ">Diversity of thought, ideas and approach is necessary for success
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Excellence</h6>
    <p class="mx-2 ">Challenge the status quo and group think, while striving to meet high expectations
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Fun</h6>
    <p class="mx-2 ">We will work hard and have fun
    </p>
    </li>
    <li class="tile p-2"> 
    <h6 class="mx-2 ">Innovate</h6>
    <p class="mx-2 ">Success is not final and failure is not terminal - take smart and ethical risk, but do
    not be afraid to fail
    </p> 
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Leadership</h6>
    <p class="mx-2 ">Lead by example - no job too small or too big
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Membership Experience
    </h6>
    <p class="mx-2 ">Do everything we can to help our Members
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Resolute</h6>
    <p class="mx-2 ">Devoted to going the extra step to ensure our company and its Membership succeed
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Leadership</h6>
    <p class="mx-2 ">Lead by example - no job too small or too big
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Membership Experience</h6>
    <p class="mx-2 ">Do everything we can to help our Members
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2">Resolute
    </h6>
    <p class="mx-2 ">Devoted to going the extra step to ensure our company and its Membership succeed
    </p>
    </li>
    <li class="tile p-2">
    <h6 class="mx-2 ">Respect
    </h6>
    <p class="mx-2 ">Respect one another, and accept there will be differences
    </p>
    </li>
    <li class="no-items">                            
    No Data to Display</li>
    </ul>  
    </div>
    </div>        
    `;     
    getFLXCulture();
  }  
   
  protected get dataVersion(): Version {
    return Version.parse('1.0');  
  }  

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [     
        {  
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
function getFLXCulture() {
  var html = "";    
  sp.web.lists
    .getByTitle("FLXCulture")
    .items.get()
    .then((items: any[]) => {
      console.log(items.length);
      if(items.length > 0){
      for (var i = 0; i < items.length; i++) {
          html += `<li class="tile p-2"><h6 class="mx-2">${items[i].Title}</h6><p class="mx-2 ">${items[i].Description}</p></li>`;   
        }}
        else {
      html = `<li class="no-items">No Data to Display </li>`
        }
      $("#tile").html(""); 
      $("#tile").html(html);
    }) 
}
  