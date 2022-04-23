import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  ISubmitActionArguments
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CourseMonitoringAdaptiveCardExtensionStrings';
import { SPHttpClient } from '@microsoft/sp-http';
import { ICourseMonitoringAdaptiveCardExtensionProps, ICourseMonitoringAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../CourseMonitoringAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<ICourseMonitoringAdaptiveCardExtensionProps, ICourseMonitoringAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: 'Reset',
        action: {
          type: 'Submit',
          parameters: {
            data: "Submit button"
          }
        }
      }
    ];
  }
  
  public onAction(action: ISubmitActionArguments): void {
    if (action.type === 'Submit') {
      this._resetProgress();
    }
  }

  private _resetProgress() {
    let body = JSON.stringify({
      'Progress': 0
    });
    
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Study Group')/items`,
      SPHttpClient.configurations.v1,
      {  
        headers: {  
          'accept': 'application/json;odata.metadata=none'
        }
      })
      .then(response => response.json())
      .then(students => {
        let students_count = students.value.length;
        for (let id = 1; id <= students_count; id++) {
          this.context.spHttpClient.post(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Study Group')/items(${id})`, 
            SPHttpClient.configurations.v1,
            {  
              headers: {  
                'accept': 'application/json;odata=nometadata',  
                'content-type': 'application/json;odata=nometadata',  
                'odata-version': '',
                'IF-MATCH': '*',  
                'X-HTTP-Method': 'MERGE'   
              },  
              body: body  
            }).catch(error => console.error(error));
          }
      })
      .catch(error => console.error(error));
  };

  public get data(): IBasicCardParameters {
    return {
      primaryText: `${this.state.bestStudent.name} - ${this.state.bestStudent.progress}%`,
      title: 'The best learner:'
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://754n8c.sharepoint.com/sites/TheBestHomeSite/Lists/Study%20Group/AllItems.aspx'
      }
    };
  }
}
