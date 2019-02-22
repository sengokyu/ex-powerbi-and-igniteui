/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
    'use strict';

    export class Visual implements IVisual {
        private host: IVisualHost;
        private submit: JQuery<HTMLButtonElement>;
        private target: JQuery<HTMLElement>;
        private table: JQuery<HTMLElement>; // ignite-ui grid
        private settings: VisualSettings;

        constructor(options: VisualConstructorOptions) {
            console.log('Visual constructor', options);
            this.host = options.host;
            this.target = jQuery(options.element);

            this.submit = jQuery('<button>') as JQuery<HTMLButtonElement>;
            this.submit.text('送信');
            this.submit.click(() => {
                const entGen = TableUtilities.entityGenerator;
                const account = 'demopbi';
                const key =
                  'Cej1y2Ba4//mjkPCcuQW6WPpNGEtFD+txBOsUq2J/QwPh6dodWp1iT3fJoqm3ggZNj9Rqp+IraT4IFLJve0y/A==';
                const table = 'demopbi';
                const tablesvc = createTableService(account, key);
                const task = {
                    PartitionKey: 'p1',
                    RowKey: entGen.String(Date.now.toString()),
                    Value: Date.now
                };
                
                tablesvc.insertEntity(table, task, (error, result, response) => {
                  if (error) {
                    console.log(error);
                  }
                });
            });

            this.table = jQuery('<table>');

            this.target.append(this.table);
            this.target.append(this.submit);

            const params: IgGrid = {
                dataSource: [
                    { Label: 1, Value: 100 },
                    { Label: 2, Value: 200 },
                    { Label: 3, Value: 300 }
                ],
                features: [
                    {
                        name: 'Responsive',
                        enableVerticalRendering: false
                    },
                    {
                        name: 'RowSelectors',
                        enableCheckBoxes: true,
                        enableRowNumbering: false,
                        enableSelectAllForPaging: true
                    },
                    {
                        name: 'Selection',
                        mode: 'row',
                        multipleSelection: true
                    }
                ]
            }


            try {
                this.table.igGrid(params);
            } catch (err) {
                console.log(err);
            }
        }

        public update(options: VisualUpdateOptions) {
            this.settings = Visual.parseSettings(
                options && options.dataViews && options.dataViews[0]
            );
            console.log('Visual update', options);
        }

        private static parseSettings(dataView: DataView): VisualSettings {
            return VisualSettings.parse(dataView) as VisualSettings;
        }

        /**
         * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
         * objects and properties you want to expose to the users in the property pane.
         *
         */
        public enumerateObjectInstances(
            options: EnumerateVisualObjectInstancesOptions
        ): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
            return VisualSettings.enumerateObjectInstances(
                this.settings || VisualSettings.getDefault(),
                options
            );
        }
    }
}
