import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
export interface ISpfx3Wp1State
{
  StateHideOrderID   : Boolean;
  StateProductItems  : IDropdownOption[];
  StateCustomerItems : IDropdownOption[];
  StateOrderItems    : IDropdownOption[];
  StateCustomerID    : string;
  StateCustomerName  : string;
  StateCustomerEmail : string;
  StateProductID     : string;
  StateProductName   : string;
  StateProductType   : string;
  StateUnitPrice     : string;
  StateUnitsSold     : string;
  StateSaleValue     : string;
  StateDate          : Date;
}
