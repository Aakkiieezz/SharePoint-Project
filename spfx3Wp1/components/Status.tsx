import * as React from 'react';
import styles from './Spfx3Wp1.module.scss';
import { ISpfx3Wp1Props } from './ISpfx3Wp1Props';
import { ISpfx3Wp1State } from './ISpfx3Wp1State';
import { sp } from "sp-pnp-js";
import { DateTimePicker, DateConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IStackTokens, Stack } from "office-ui-fabric-react/lib/Stack";
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import "@pnp/sp/lists";
import "@pnp/sp/items";

var CustomerNamesArray   : IDropdownOption[] = [];
var ProductNamesArray    : IDropdownOption[] = [];
var OrderItemsArray      : IDropdownOption[] = [];
var SelectedCustomerName : string;
var SelectedProductName  : string;
var SelectedOrderID;

export default class Status extends React.Component<ISpfx3Wp1Props, ISpfx3Wp1State>
{
  constructor(props: ISpfx3Wp1Props, state: ISpfx3Wp1State)
  {
    super(props);
    this.state =
    {
      StateHideOrderID   : false,
      StateProductItems  : [],
      StateCustomerItems : [],
      StateOrderItems    : [],
      StateProductType   : '',
      StateDate          : new Date(),
      StateUnitPrice     : '',
      StateUnitsSold     : '',
      StateSaleValue     : '',
      StateCustomerID    : '',
      StateProductID     : '',
      StateCustomerName  : '',
      StateCustomerEmail : '',
      StateProductName   : ''
    };
    this.handleChangeCustomer = this.handleChangeCustomer.bind(this);
    this.handleChangeProduct  = this.handleChangeProduct.bind(this);
    this.handleChangeOrder    = this.handleChangeOrder.bind(this);
    this.handleUnitChange     = this.handleUnitChange.bind(this);
    this.deleteItem           = this.deleteItem.bind(this);
    this.resetForm            = this.resetForm.bind(this);
    this.editItem             = this.editItem.bind(this);
    this.addItem              = this.addItem.bind(this);    
  }

  public async componentDidMount(): Promise<void>
  {
    sp.web.lists.getByTitle("Customers").items.select('Title').get().then((data)=>
    {
      for(var k in data)
        CustomerNamesArray.push({ key: data[k].Title, text: data[k].Title });
      this.setState({ StateCustomerItems : CustomerNamesArray });
      // return CustomerNamesArray;
    });
    
    sp.web.lists.getByTitle("Products").items.select('Title').get().then((data)=>
    {
      for(var k in data)
        ProductNamesArray.push({ key: data[k].Title, text: data[k].Title });
      this.setState({ StateProductItems : ProductNamesArray });
      // return ProductNamesArray;
    });
    
    sp.web.lists.getByTitle("Orders").items.select('ID').get().then((data)=>
    {
      for(var k in data)
        OrderItemsArray.push({ key: data[k].ID, text: data[k].ID });
      this.setState({ StateOrderItems : OrderItemsArray });
      // return OrderItemsArray;
    });
  }

  async handleChangeCustomer(event): Promise<void>
  {
    try
    {
      SelectedCustomerName = event.key;
      let items = await sp.web.lists.getByTitle("Customers").items.getPaged();
      for(let i=0; i<items.results.length; i++)
        if(items.results[i].Title == SelectedCustomerName)
        {
          this.setState({ StateCustomerID    : items.results[i].ID                   });
          this.setState({ StateCustomerName  : items.results[i].Title                });
          this.setState({ StateCustomerEmail : items.results[i].CustomerEmailAddress });
          break;
        }
    }
    catch(error)
    {
      console.error(error);
    }
  }

  async handleChangeProduct(event): Promise<void>
  {
    try
    {
      SelectedProductName = event.key;
      let productsArray = await sp.web.lists.getByTitle("Products").items.getPaged();
      for(let i=0; i<productsArray.results.length; i++)
        if(productsArray.results[i].Title == SelectedProductName)
        {
          this.setState({ StateProductID   : productsArray.results[i].ID                          });
          this.setState({ StateProductName : productsArray.results[i].Title                       });
          this.setState({ StateUnitPrice   : productsArray.results[i].ProductUnitPrice            });
          this.setState({ StateProductType : productsArray.results[i].ProductType                 });
          this.setState({ StateDate        : new Date(productsArray.results[i].ProductExpiryDate) });
          {
            var units     : number = parseInt(this.state.StateUnitsSold);
            var unitPrice : number = parseInt(this.state.StateUnitPrice);
            var calculate = units * unitPrice;
            this.setState({ StateSaleValue: calculate.toString() });
          }
          break;
        }
    }
    catch(error)
    {
      console.error(error);
    }
  }

  async handleChangeOrder(event): Promise<void>
  {
    try
    {
      SelectedOrderID = event.key;
      
      let items = await sp.web.lists.getByTitle("Orders").items.getPaged();
      for(let i=0; i<items.results.length; i++)
        if(items.results[i].ID == SelectedOrderID)
        {
          this.setState({ StateCustomerName  : items.results[i].CustomerName   });
          this.setState({ StateProductName   : items.results[i].ProductName    });
          this.setState({ StateProductID     : items.results[i].ProductID      });
          this.setState({ StateCustomerID    : items.results[i].CustomersID    });
          this.setState({ StateCustomerEmail : items.results[i].CustomersEmail });
          this.setState({ StateUnitPrice     : items.results[i].UnitPrice      });
          this.setState({ StateUnitsSold     : items.results[i].UnitsSold      });
          this.setState({ StateSaleValue     : items.results[i].SaleValue      });
          break;
        }
      
      let productsArray = await sp.web.lists.getByTitle("Products").items.getPaged();
      for(let i=0; i<productsArray.results.length; i++)
        if(productsArray.results[i].ID == this.state.StateProductID)
        {
          this.setState({ StateProductType : productsArray.results[i].ProductType });
          this.setState({ StateDate        : new Date(productsArray.results[i].ProductExpiryDate) });
          break;
        }
    }
    catch(error)
    {
      console.error(error);
    }
  }

  public handleUnitChange = (event) =>
  {
    this.setState({ StateUnitsSold: event.target.value.toString() });
    var units     : number = parseInt(event.target.value);
    var unitPrice : number = parseInt(this.state.StateUnitPrice);
    var calculate = units * unitPrice;
    this.setState({ StateSaleValue: calculate.toString() });
    return event;
  }

  async addItem(): Promise<void>
  {
    try
    {
      if(this.state.StateCustomerName == "")
        alert("Please Select Customer Name From Dropdown");
      else if(this.state.StateProductName == "")
        alert("Please Select Product Name From Dropdown");
      else if(this.state.StateUnitsSold == "" || this.state.StateUnitsSold == "0")
        alert("Please Enter Number Of Units (greater than 0)");
      else if(Number(this.state.StateUnitsSold) !== parseInt(this.state.StateUnitsSold) && Number(this.state.StateUnitsSold) % 1 !== 0)
        alert("Please enter No. of Units as integer value");
      else
      {
        await sp.web.lists.getByTitle("Orders").items.add(
        {
          CustomersID    : parseInt(this.state.StateCustomerID),
          CustomerName   : this.state.StateCustomerName,
          CustomersEmail : this.state.StateCustomerEmail,
          ProductID      : parseInt(this.state.StateProductID),
          ProductName    : this.state.StateProductName,
          ProductType    : this.state.StateProductType,
          UnitPrice      : parseInt(this.state.StateUnitPrice),
          UnitsSold      : parseInt(this.state.StateUnitsSold),
          SaleValue      : parseInt(this.state.StateSaleValue)
        });
        var latestOrderID = OrderItemsArray[OrderItemsArray.length - 1].key;
        if(typeof(latestOrderID) == 'number')
        {
          let id = latestOrderID + 1;
          alert("New Order ID "+id+" added Successfully");
        }
        location.reload();
      }
    }
    catch(error)
    {
      console.error(error);
      alert(error.message);
    }
  }

  private resetForm(): void
  {
    SelectedCustomerName = '';
    SelectedProductName  = '';
    SelectedOrderID      = '';
    this.setState(
    {
      StateCustomerID    : '',
      StateCustomerName  : '',
      StateCustomerEmail : '',
      StateProductID     : '',
      StateProductName   : '',
      StateProductType   : '',
      StateUnitPrice     : '',
      StateUnitsSold     : '',
      StateSaleValue     : '',
      StateDate          : new Date()     
    });
  }

  async editItem(): Promise<void>
  {
    try
    {
      this.setState({ StateHideOrderID: true });
      await sp.web.lists.getByTitle("Orders").items.getById(SelectedOrderID).update(
      {
        CustomersID  : parseInt(this.state.StateCustomerID),
        CustomerName : this.state.StateCustomerName,
        ProductID    : parseInt(this.state.StateProductID),
        ProductName  : this.state.StateProductName,
        ProductType  : this.state.StateProductType,
        UnitPrice    : parseInt(this.state.StateUnitPrice),
        UnitsSold    : parseInt(this.state.StateUnitsSold),
        SaleValue    : parseInt(this.state.StateSaleValue)
      });
      alert("Order ID "+SelectedOrderID+" updated Successfully");
      location.reload();
    }
    catch(error)
    {
      console.error(error);
    }
  }

  async deleteItem(): Promise<void>
  {
    try
    {
      this.setState({ StateHideOrderID: true });
      await sp.web.lists.getByTitle("Orders").items.getById(SelectedOrderID).delete();
      alert("Order ID "+SelectedOrderID+" deleted Successfully");
      location.reload();
    }
    catch(error)
    {
      console.error(error);
    }
  }

  public render(): React.ReactElement<ISpfx3Wp1Props>
  {
    const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 452 } };
    const stackTokens: IStackTokens = { childrenGap: 20 };
    return (
      <div className = { styles.spfx3Wp1 }>
        <div>
          <div>
            <header>
              <h1> <img src = { require('./images/logo.png') } alt="logo" width="55" height="50" />&nbsp; &nbsp;ORDER FORM</h1>
            </header>
          </div>
          <div className = { styles.row }>
            <div className = { styles.column }>
              
              <Stack
              tokens = { stackTokens }>
                <Dropdown
                required      = { true }
                  placeholder = "Select a Customers"
                  label       = "Customer Name"
                  selectedKey = { this.state.StateCustomerName }
                  options     = { this.state.StateCustomerItems }
                  styles      = { dropdownStyles }
                  onChanged   = { this.handleChangeCustomer } />
              </Stack>
              
              <Stack
                tokens = {stackTokens}>
                <Dropdown
                  required    = {true}
                  placeholder = "Select a Product"
                  label       = "Product Name"
                  selectedKey = { this.state.StateProductName }
                  options     = { this.state.StateProductItems }
                  styles      = { dropdownStyles }
                  onChanged   = { this.handleChangeProduct } />
              </Stack>
              
              <TextField
                required       = {true}
                placeholder    = "Product Type"
                label          = "Product Type"
                value          = {this.state.StateProductType}
                onChange       = {e => { this.setState({ StateProductType: this.state.StateProductType }); }} />
              
              <DateTimePicker
                label          = "Product Expiry"
                dateConvention = {DateConvention.Date}
                value          = {this.state.StateDate} />
              
              <TextField
                required    = {true}
                placeholder = "Product Unit Price"
                label       = "Product Unit Price"
                type        = "number"
                value       = {this.state.StateUnitPrice}
                onChange    = {e => { this.setState({ StateUnitPrice : this.state.StateUnitPrice }); }} />
              
              <TextField
                required    = {true}
                placeholder = "Enter No. of Units > 0"
                label       = "Number of Units"
                type        = "number"
                value       = {this.state.StateUnitsSold}
                onChange    = {this.handleUnitChange} />
              
              <TextField
                required    = {true}
                placeholder = "Sale Value"
                label       = "Sale Value"
                type        = "number"
                value       = {this.state.StateSaleValue}
                onChange    = {e => { this.setState({ StateSaleValue: this.state.StateSaleValue }); }} />
                
              <br></br>
              
              <button onClick={this.addItem}    className={styles.buttonAdd}> ADD   </button> &nbsp; &nbsp;
              
              <button onClick={this.resetForm}  className={styles.buttonReset}> RESET  </button>
              
              <Stack tokens={stackTokens}>
                {
                  this.state.StateHideOrderID ?
                    <Dropdown
                      required    = { true }
                      placeholder = "Select an Order ID to Edit/Delete"
                      label       = "Order ID"
                      selectedKey = { SelectedOrderID }
                      options     = { this.state.StateOrderItems }
                      styles      = { dropdownStyles }
                      onChanged   = { this.handleChangeOrder }/>
                    : null
                }
              </Stack> <br></br>
              
              <button onClick={this.editItem}   className={styles.buttonEdit}> EDIT   </button> &nbsp; &nbsp;
              
              <button onClick={this.deleteItem} className={styles.buttonDelete}> DELETE </button>
            </div>
          </div>
          <div>
            <footer>
              <section>
                <div>
                  <h4>©2021 shop.adidas.co.in | Powered By : adi Sports (India) Pvt. Ltd.</h4>
                </div>
              </section>
            </footer>
          </div>
        
        </div>
      </div>
    );
  }
}