import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneChoiceGroup, PropertyPaneSlider, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/lists';
import '@pnp/sp/security/list';
import '@pnp/sp/security/web';

import {
    SharepointOnlineSiteScanner,
    ISharepointOnlineSiteScannerProps,
    UITheme
} from './components/SharepointOnlineSiteScanner';
import { LoggerService, LogLevel } from './services/LoggerService';

import * as strings from 'SharepointOnlineSiteScannerWebPartStrings';

export interface ISharepointOnlineSiteScannerWebPartProps {
    internalDomains: string;
    uiTheme: UITheme;
    riskWeightOversharing: number;
    riskWeightExternalUser: number;
    riskWeightFileSharing: number;
    riskWeightGovernance: number;
}

export default class SharepointOnlineSiteScannerWebPart extends BaseClientSideWebPart<ISharepointOnlineSiteScannerWebPartProps> {
    private _sp!: SPFI;

    protected async onInit(): Promise<void> {
        await super.onInit();

        // Initialize PnPjs with SPFx context
        this._sp = spfi().using(SPFx(this.context));

        // Set logger level based on environment
        if (DEBUG) {
            LoggerService.setLevel(LogLevel.Debug);
            LoggerService.debug('WebPart', 'Debug mode enabled');
        } else {
            LoggerService.setLevel(LogLevel.Warning);
        }

        LoggerService.info('WebPart', `Initialized for site: ${this.context.pageContext.web.absoluteUrl}`);
    }

    public render(): void {
        const element: React.ReactElement<ISharepointOnlineSiteScannerProps> = React.createElement(
            SharepointOnlineSiteScanner,
            {
                sp: this._sp,
                context: this.context,
                internalDomains: this.properties.internalDomains || '',
                siteUrl: this.context.pageContext.web.absoluteUrl,
                uiTheme: this.properties.uiTheme || 'classic',
                customCaps: {
                    OVERSHARING_RISK: this.properties.riskWeightOversharing ?? 40,
                    EXTERNAL_USER_RISK: this.properties.riskWeightExternalUser ?? 30,
                    FILE_SHARING_RISK: this.properties.riskWeightFileSharing ?? 20,
                    GOVERNANCE_HYGIENE: this.properties.riskWeightGovernance ?? 10
                }
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('internalDomains', {
                                    label: strings.InternalDomainsFieldLabel,
                                    description: strings.InternalDomainsDescription,
                                    multiline: true,
                                    rows: 3,
                                    placeholder: 'contoso.com, contoso.onmicrosoft.com',
                                }),
                                PropertyPaneChoiceGroup('uiTheme', {
                                    label: 'UI Theme',
                                    options: [
                                        { key: 'classic', text: 'Classic (Data Dense)' },
                                        { key: 'modern', text: 'Modern (Clean & Simple)' },
                                        { key: 'compact', text: 'Compact (Cards & Tables)' },
                                        { key: 'unified', text: 'Unified Report (Scrollable)' }
                                    ]
                                })
                            ],
                        },
                        {
                            groupName: 'Risk Scoring Weights (Must equal 100)',
                            groupFields: [
                                PropertyPaneSlider('riskWeightOversharing', {
                                    label: 'Copilot / Oversharing Risk Weight',
                                    min: 0,
                                    max: 100,
                                    step: 5,
                                    value: this.properties.riskWeightOversharing ?? 40
                                }),
                                PropertyPaneSlider('riskWeightExternalUser', {
                                    label: 'External User Risk Weight',
                                    min: 0,
                                    max: 100,
                                    step: 5,
                                    value: this.properties.riskWeightExternalUser ?? 30
                                }),
                                PropertyPaneSlider('riskWeightFileSharing', {
                                    label: 'File Sharing Risk Weight',
                                    min: 0,
                                    max: 100,
                                    step: 5,
                                    value: this.properties.riskWeightFileSharing ?? 20
                                }),
                                PropertyPaneSlider('riskWeightGovernance', {
                                    label: 'Governance Hygiene Risk Weight',
                                    min: 0,
                                    max: 100,
                                    step: 5,
                                    value: this.properties.riskWeightGovernance ?? 10
                                })
                            ]
                        }
                    ],
                },
            ],
        };
    }

    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
        const p = this.properties;
        const total = (p.riskWeightOversharing ?? 40) + 
                      (p.riskWeightExternalUser ?? 30) + 
                      (p.riskWeightFileSharing ?? 20) + 
                      (p.riskWeightGovernance ?? 10);
                      
        if (total !== 100) {
            this.context.propertyPane.refresh();
        }
    }
    
    protected async onPropertyPaneConfigurationStart(): Promise<void> {
        // Force validation on load if needed
    }
}
