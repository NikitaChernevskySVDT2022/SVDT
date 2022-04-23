import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { CourseMonitoringPropertyPane } from './CourseMonitoringPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ICourseMonitoringAdaptiveCardExtensionProps {
  title: string;
}

export interface ICourseMonitoringAdaptiveCardExtensionState {
  bestStudent: IBestStudent | undefined;
}

const CARD_VIEW_REGISTRY_ID: string = 'CourseMonitoring_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'CourseMonitoring_QUICK_VIEW';

export default class CourseMonitoringAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICourseMonitoringAdaptiveCardExtensionProps,
  ICourseMonitoringAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CourseMonitoringPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { 
      bestStudent: undefined
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchBestStudent();
  }

  private _fetchBestStudent(): Promise<void> {
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Study Group')/items?&$select=Title,Progress`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'accept': 'application/json;odata.metadata=none'
        }
      })
      .then(response => response.json())
      .then(students => {
        const students_array = students.value;
        const bestStudent = students_array.reduce(function(prev, current) {
          return (prev.Progress >= current.Progress) ? prev : current
        });

        this.setState({
          bestStudent: {
            name: bestStudent.Title,
            progress: Math.round(bestStudent.Progress * 100)
          }
        });
      })
      .catch(error => console.error(error));
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'CourseMonitoring-property-pane'*/
      './CourseMonitoringPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.CourseMonitoringPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
export interface IBestStudent {
  name: string;
  progress: number;
}