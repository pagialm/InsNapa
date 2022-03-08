import { Dropdown, IDropdownOption } from "office-ui-fabric-react";
import * as React from "react";
import { IFilteredDropdownProps } from "./IFilteredDropdown";
import { SPHttpClient } from "@microsoft/sp-http";
import { IFieldSP } from "./IFieldSP";

export interface IFilteredDropdownState {
  options: IDropdownOption[];
}
export interface IProduct {
  Business: string;
  Product: string;
  Product_x0020_Area: string;
}

export default class FilteredDropdown extends React.Component<
  IFilteredDropdownProps,
  IFilteredDropdownState
> {
  constructor(props: IFilteredDropdownProps) {
    super(props);
    this.state = {
      options: [],
    };
  }
  public componentDidMount() {
    this._getItems(
      this.props.listname,
      this.props.field1,
      this.props.field2,
      this.props.field3
    ).then((items) => {
      let optionsArray: IDropdownOption[] = [];
      console.log(items);
      items.forEach((item) => {
        if (this.props.field3) {
          const product = item as IProduct;
          optionsArray.push({ key: product.Product, text: product.Product });
        } else if (this.props.field2) {
          const product = item as IProduct;
          optionsArray.push({ key: product.Business, text: product.Business });
        } else {
          const product = item as IProduct;
          optionsArray.push({
            key: product.Product_x0020_Area,
            text: product.Product_x0020_Area,
          });
        }
      });
      this.setState({ options: optionsArray });
    });
  }
  private _getItems(
    listName: string,
    field1: IFieldSP,
    field2?: IFieldSP,
    field3?: IFieldSP
  ): Promise<any> {
    let url: string =
      this.props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items?$filter=" +
      field1.name +
      " eq '" +
      field1.value +
      "'";
    if (field2) url += " and " + field2.name + " eq '" + field2.value + "'";
    if (field3) url += " and " + field3.name + " eq '" + field3.value + "'";
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<any>;
  }
  public render(): React.ReactElement<IFilteredDropdownProps> {
    return (
      <div>
        <Dropdown options={this.state.options} label={this.props.label} />
      </div>
    );
  }
}
