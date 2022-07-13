import * as React from "react";
import { IMfrpSalesFormProps } from "./IMfrpSalesFormProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IMfrpSalesFormState } from "./IMfrpSalesFormState";
export default class MfrpSalesForm extends React.Component<IMfrpSalesFormProps, IMfrpSalesFormState, {}> {
    render(): React.ReactElement<IMfrpSalesFormProps>;
    constructor(props: IMfrpSalesFormProps, state: IMfrpSalesFormState);
    componentDidMount(): void;
    private readOrderList;
    private Dropdowns;
    AutoPopulate(): Promise<void>;
    addItems(): Promise<void>;
    private resetItems;
    private editItem;
    private updateItem;
    private deleteItem;
}
//# sourceMappingURL=MfrpSalesForm.d.ts.map