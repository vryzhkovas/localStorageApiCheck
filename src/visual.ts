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
"use strict";

import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualLocalStorageV2Service = powerbi.extensibility.IVisualLocalStorageV2Service;
import PrivilegeStatus = powerbi.PrivilegeStatus; 

export class Visual implements IVisual {
    private target: HTMLElement;
    private get_input: HTMLInputElement;
    private get_result_text: Text;
    private status_result_text: Text;
    private set_name_input: HTMLInputElement;
    private set_value_input: HTMLInputElement;
    private remove_input: HTMLInputElement;
    private get_p: HTMLElement;
    private storageV2Service: IVisualLocalStorageV2Service; 

    constructor(options: VisualConstructorOptions) {
        this.storageV2Service = options.host.storageV2Service;
        this.target = options.element;

        if (document) {
            const status_container = document.createElement("div");
            const status_p: HTMLElement = document.createElement("p");
            const status_button: HTMLElement = document.createElement("button");
            status_button.setAttribute("id", "statusButton");
            status_button.textContent = "Check the status";
            status_button.onclick = () => this.onStatusButtonClick();
            status_p.appendChild(status_button);

            const status_result_p: HTMLElement = document.createElement("p");
            status_result_p.appendChild(document.createTextNode("Local storage availability:"));
            const status_result_em: HTMLElement = document.createElement("em");
            this.status_result_text = document.createTextNode("");
            status_result_em.appendChild(this.status_result_text);
            status_result_p.appendChild(status_result_em);

            status_container.appendChild(status_p);
            status_container.appendChild(status_result_p);
            this.target.appendChild(status_container);

            const set_container = document.createElement("div");
            const set_p: HTMLElement = document.createElement("p");
            this.set_name_input = document.createElement("input");
            this.set_name_input.setAttribute("type", "text");
            this.set_name_input.setAttribute("placeholder", "Name");
            this.set_value_input = document.createElement("input");
            this.set_value_input.setAttribute("type", "text");
            this.set_value_input.setAttribute("placeholder", "Value");
            const set_button: HTMLElement = document.createElement("button");
            set_button.textContent = "Set";
            set_button.onclick = () => this.onSetButtonClick();
            set_p.appendChild(this.set_name_input);
            set_p.appendChild(this.set_value_input);
            set_p.appendChild(set_button);

            set_container.appendChild(set_p);
            this.target.appendChild(set_container);

            const get_container = document.createElement("div");
            this.get_p = document.createElement("p");
            this.get_input = document.createElement("input");
            this.get_input.setAttribute("type", "text");
            const get_button: HTMLElement = document.createElement("button");
            get_button.textContent = "Get";
            get_button.onclick = () => this.onGetButtonClick();
            this.get_p.appendChild(this.get_input);
            this.get_p.appendChild(get_button);

            const get_result_p: HTMLElement = document.createElement("p");
            get_result_p.appendChild(document.createTextNode("Result:"));
            const get_result_em: HTMLElement = document.createElement("em");
            this.get_result_text = document.createTextNode("");
            get_result_em.appendChild(this.get_result_text);
            get_result_p.appendChild(get_result_em);

            get_container.appendChild(this.get_p);
            get_container.appendChild(get_result_p);
            this.target.appendChild(get_container);

            const remove_container = document.createElement("div");
            const remove_p: HTMLElement = document.createElement("p");
            this.remove_input = document.createElement("input");
            this.remove_input.setAttribute("type", "text");
            const remove_button: HTMLElement = document.createElement("button");
            remove_button.textContent = "Remove";
            remove_button.onclick = () => this.onRemoveButtonClick();
            remove_p.appendChild(this.remove_input);
            remove_p.appendChild(remove_button);

            remove_container.appendChild(remove_p);
            this.target.appendChild(remove_container);
        }
    }
    public async onStatusButtonClick(): Promise<void> {
        this.status_result_text.textContent = (await this.storageV2Service.status()).toString();
    }

    public async onGetButtonClick(): Promise<void> {
        try { 
            console.log(`Get ${this.get_input.value}`);
            let status: PrivilegeStatus = await this.storageV2Service.status(); 
            if (status === PrivilegeStatus.Allowed) { 
                this.get_result_text.textContent = await this.storageV2Service.get(this.get_input.value);
                this.removeErrorMessage();
            }
        }
        catch {
            let errorMessage = this.get_p.querySelector(".errorMessage");
            if (!errorMessage) {
                this.get_p.appendChild(this.createErrorMessage("Error: wrong name"));
                this.get_result_text.textContent = "";
            }
        }
    }

    public async onSetButtonClick(): Promise<void> {
        console.log(`Set key:${this.set_name_input.value} value:${this.set_value_input.value}`);
        let status: PrivilegeStatus = await this.storageV2Service.status(); 
        if (status === PrivilegeStatus.Allowed) {
            await this.storageV2Service.set(this.set_name_input.value, this.set_value_input.value);
        }
    }

    public async onRemoveButtonClick(): Promise<void> {
        console.log(`Remove ${this.remove_input.value}`);
        let status: PrivilegeStatus = await this.storageV2Service.status(); 
        if (status === PrivilegeStatus.Allowed) {
            await this.storageV2Service.remove(this.remove_input.value);
        }
    }

    public removeErrorMessage() {
        let errorMessage = this.get_p.querySelector(".errorMessage");
        if (errorMessage) {
            errorMessage.remove();
        }
    }

    public createErrorMessage(message: string): HTMLElement {
        let errorMessage = document.createElement("div");
        errorMessage.className = "errorMessage";
        errorMessage.appendChild(document.createTextNode(message));
        let crossButton = document.createElement("button");
        crossButton.className = "closeBtn";
        crossButton.textContent = "X";
        crossButton.onclick = () => this.removeErrorMessage();
        errorMessage.appendChild(crossButton);
        return errorMessage;
    }

    public update(options: VisualUpdateOptions): void {
    }
}