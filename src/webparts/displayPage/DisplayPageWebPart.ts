import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './DisplayPageWebPart.module.scss';
import * as strings from 'DisplayPageWebPartStrings';

import * as $ from 'jquery';
import axios from 'axios';

export interface IDisplayPageWebPartProps {
  description: string;
}

export default class DisplayPageWebPart extends BaseClientSideWebPart<IDisplayPageWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="ms-Fabric" dir="ltr">
      <div id="projectDetails"></div>
    
      <div id="projectObjective" class="ProjectObjective"></div>

      <div id="executiveStatus" class="ExecutiveStatus"></div>

      <div class="${styles.RAGStatus}">
          <h3>RAG Status</h3>
          <table>
          <thead>
              <tr>
                  <th class="ms-textAlignCenter">Project Cost</th>
                  <th class="ms-textAlignCenter">Project Scope</th>
                  <th class="ms-textAlignCenter">Project Resource</th>
                  <th class="ms-textAlignCenter">Project Schedule</th>
              </tr>
          </thead>
          <tbody>
              <tr>
                  <td id="colorCost"></td>
                  <td id="colorScope"></td>
                  <td id="colorResource"></td>
                  <td id="colorSchedule"></td>
              </tr>
              <tr>
                  <td id="iconCost" class="ms-textAlignCenter"></td>
                  <td id="iconScope" class="ms-textAlignCenter"></td>
                  <td id="iconResource" class="ms-textAlignCenter"></td>
                  <td id="iconSchedule" class="ms-textAlignCenter"></td>
              </tr>
          </tbody>
          </table>
      </div>

          <div class="KeyMilestones">
          <h3>Key Milestones</h3>
              <table>
                <tbody id="keyMilestones"></tbody>
              </table>
          </div>

          <div class="CompletedThisPeriod">
            <h3>Completed This Period</h3>
            <table>
              <thead class="WithoutBorder">
                <tr>
                <th>Task</th>
                <th class="narrow">Date</th>
                </tr>
              </thead>
              <tbody id="completedThisPeriod" class="WithoutBorder"></tbody>
            </table>
          </div>

          <div class="FocusForNextPeriod">
          <h3>Focus for Next Period</h3>
            <table>
            <thead>
              <tr>
                <th>Task</th>
                <th class="narrow">Date</th>
              </tr>
            </thead>
            <tbody id="focusForNextPeriod"></tbody>
            </table>
          </div>
    
          <div class="clear"></div>

          <div class="ProjectRisk">
          <h3>Project Risk</h3>
          <table>
            <thead>
              <tr>
                <th>Risk Ref</th>
                <th>Description</th>
                <th class="narrow">Gross Risk</th>
                <th>Mitigation</th>
                <th class="narrow">Net Risk</th>
                <th>Risk Owner</th>
              </tr>
            </thead>
            <tbody id="projectRisks"></tbody>
          </table>
          </div>

          <div class="ProjectIssue">
          <h3>Project Issue</h3>
          <table>
            <thead>
              <tr>
                <th>Issue Ref</th>
                <th>Description</th>
                <th>Escalation</th>
                <th>Actions</th>
                <th>Action Due</th>
                <th>Issue Owner</th>
              </tr>
            </thead>
            <tbody id="projectIssues"></tbody>
          </table>
      </div>
    </div>`;

    require('./my-script.js');
    $("nothingSelected").show();
    axios('https://jsonplaceholder.typicode.com/todos/1');
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
