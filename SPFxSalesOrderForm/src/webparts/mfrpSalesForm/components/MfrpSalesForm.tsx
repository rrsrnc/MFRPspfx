import * as React from "react";
import styles from "./MfrpSalesForm.module.scss";
import { IMfrpSalesFormProps } from "./IMfrpSalesFormProps";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {IMfrpSalesFormState,IProductName,ICustomerName} from "./IMfrpSalesFormState";
import * as $ from "jquery";
export default class MfrpSalesForm extends React.Component<IMfrpSalesFormProps,IMfrpSalesFormState,{}> {
  public render(): React.ReactElement<IMfrpSalesFormProps> {
    return (
      <div className={styles.mfrpSalesForm}>
        <div className={styles.container}>
          <div className={styles.row}>
            <h2 className={styles.title}>Sales Order Form</h2>
            <div>
              ===========================================================================================
            </div>
          </div>
          <div className={styles.row}>
            <label className={styles.label} htmlFor="customerName">
              Customer Name
            </label>
            <select className={styles.input} id="customerName" placeholder="Customer Name"></select>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="productName">
              Product Name
            </label>
            <select className={styles.input} id="productName" placeholder="Product Name" ></select>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="productType">
              Product Type
            </label>
            <input className={styles.input} id="productType"></input>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="productExpiryDate">
              Product Expiry date
            </label>
            <input className={styles.input} id="productExpiryDate"></input>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="productUnitPrice">
              Product Unit Price
            </label>
            <input className={styles.input} id="productUnitPrice"></input>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="noOfUnits">
              Number of Units
            </label>
            <input className={styles.input} id="noOfUnits"></input>
            <br></br>
            <br></br>
            <label className={styles.label} htmlFor="salesValue">
              Sales Value
            </label>
            <input className={styles.input} id="salesValue"></input>
            <br></br>
            <br></br>
            <button className={styles.add} id="add" onClick={() => this.addItems()}>
              Add
            </button>
            <button className={styles.update} id="update" onClick={() => this.updateItem()}>
              Update
            </button>
            <button className={styles.reset} id="reset" onClick={() => this.resetItems()}>
              Reset
            </button>
            <br />
            <h2 className={styles.title}>Edit OR Delete Order</h2>
            <div>
              ===========================================================================================
            </div>
            <br></br>
            <label className={styles.label} htmlFor="orderId">
              Enter Order Id
            </label>
            <input className={styles.input} id="orderId" placeholder="Format=OID-digits"></input>
            <button className={styles.edit} id="edit" onClick={() => this.editItem()}>
              Edit
            </button>
            <button className={styles.delete} id="delete" onClick={() => this.deleteItem()}>
              Delete
            </button>
            <h2 className={styles.title}>Orders List Items</h2>
            <div className={styles.row}>
              <table id="orderList"></table>
            </div>
          </div>
        </div>
      </div>
    );
  }
  //Constructor
  constructor(props: IMfrpSalesFormProps, state: IMfrpSalesFormState) {
    super(props);
    sp.setup({
      spfxContext: this.props.spcontext,
    });
    state = {
      items: [],
      Products: [],
    };
  }

  public componentDidMount() {
    this.readOrderList();
    this.Dropdowns();
    
  }


  // Read Order List Items
  private async readOrderList() {
   try{
       var orderItems: any[] = await sp.web.lists.getByTitle("Orders").items.getAll();
       var orderHtml = `<tr><th>Orders Id</th>
                  <th>Customer Id</th>  
                  <th>Product Id</th>
                  <th>Unit Sold</th>   
                  <th>Unit Price</th>   
                  <th>Sales Value</th>
                  <th>Order Status</th>
                  </tr>`;

        orderItems.forEach((element) => {
        orderHtml += `<tr><td>${element.OrdersID}</td><td>${element.CustomerID}</td>
                  <td>${element.ProductID}</td> <td>${element.UnitPrice}</td>
                  <td>${element.UnitSold}</td><td>${element.SalesValue}</td>
                  <td>${element.OrderStatus}</td></tr>`;
         });
        $("#orderList").html(orderHtml);
      }
   catch(error){
     alert(error.message);
    }  
  }


  // DropDowns
  private async Dropdowns() {
   try{
        var items: ICustomerName[] = [];

        var customerItems: any[] = await sp.web.lists
          .getByTitle("Customers")
          .items.getAll();

        customerItems.forEach((element) => {
          items.push({ customerName: element.CustomerName });
        });
        this.setState({ items: items });
        var html = "<option>Select Customer Name</option>";

        items.forEach(function (item) {
          html += `<option>${item.customerName}</option>`;
        });
        $("#customerName").append(html);

        // Product Details
        var products: IProductName[] = [];
        var productItems: any[] = await sp.web.lists
          .getByTitle("Products")
          .items.getAll();

        this.setState({ Products: productItems });
        
        productItems.forEach((element) => {
          products.push({ productName: element.ProductName });
        });

        var prohtml = "<option>Select Product Name</option>";
        products.forEach(function (item) {
          prohtml += `<option>${item.productName}</option>`;
        });

        $("#productName").append(prohtml);
        this.AutoPopulate();
    }
  catch(error){
    alert(error.message);
  }
 }

  //Auto populate
  public async AutoPopulate() {
   try{
    var productItems: any[] = await sp.web.lists.getByTitle("Products").items.getAll();
    
    $("#productName").on("change", function () {
      var proName = $("#productName option:selected").val();
    productItems.forEach((element) => {
        if (element.ProductName == proName) {
          // alert("found")
          var date: string = element.ProductExpiryDate;
          $("#productType").val(element.ProductType);
          $("#productUnitPrice").val(element.ProductUnitPrice);
          $("#productExpiryDate").val(date.substr(0, 10));
        }
      //  else
      //  alert("Product Not found")
      });
      if($('#noOfUnits').val()!="")
      {
          let unitPrice:number=parseFloat($('#productUnitPrice').val().toString());
          let units:number=parseFloat($('#noOfUnits').val().toString());
          let sales:number=unitPrice*units;
          $('#salesValue').val(sales);
      }
    });
          
    $("#noOfUnits").on("change", function () {
      let units: number = parseFloat($("#noOfUnits").val().toString());
      let price: number = parseFloat($("#productUnitPrice").val().toString());

      if ($("#noOfUnits").val() == "") 
      {
        alert("Please enter no of units");
      } 
      else 
      {
        let sales = units * price;
        $("#salesValue").val(sales);
      }
    });
  }
  catch(error){
    alert(error.message);
  }
 }

  //  Add items to order list
  public async addItems() {
    try{
        var title: string='';
        var orderId: string='';
        var custId: string='';
        var proId: string='';
        var unitSold: string = $("#noOfUnits").val().toString();
        var unitPrice = $("#productUnitPrice").val();
        var salesValue = $("#salesValue").val();
        var id: number=0;
        var custName = $("#customerName").val();
        var proName = $("#productName").val();
      
        if (custName == "Select Customer Name")
          alert("Please Select Customer Name");
        else if (proName == "Select Product Name")
          alert("Please Select Product Name");
        else if (unitSold == "") 
        alert("please input number of units sold");
        else {
          var orderListUrl = sp.web.lists.getByTitle("Orders").items;
          var orderItems: any[] = await orderListUrl.getAll();
          var products = await sp.web.lists.getByTitle("Products").items.getAll();
          var customers = await sp.web.lists.getByTitle("Customers").items.getAll();
          if (orderItems.length == 0) 
          {
            id = 0;
            title = "1";
          } 
          else 
          {
            orderItems.forEach((element) => {
              id = element.ID;
              title = element.Title;
            });
          }

          customers.forEach((element) => {
            if (element.CustomerName == custName) 
            {
              custId = element.CustomerID;
            }
          });

          products.forEach((element) => {
            if (element.ProductName == proName) 
            {
              proId = element.ProductID;
            }
          });

          var titleNum = parseInt(title);
          title = (titleNum += 1).toString();
          orderId = "OID-" + (id += 1);

          orderListUrl
            .add({
              Title: title,
              OrdersID: orderId,
              CustomerID: custId,
              ProductID: proId,
              UnitSold: unitSold,
              UnitPrice: unitPrice,
              SalesValue: salesValue,
              OrderStatus: "Approved",
            })
            .then(() => alert(`New Item with orderId ${orderId} added successfully to Order List`))
            .then(() => this.readOrderList())
            .catch((e) => alert("Error" + e.message));
          this.resetItems();
      }
    }
        catch(error){
          alert(error.message);
        }
  }

  // Reset Items in input Fields
  private resetItems() {
    try{
        $("#customerName").val("Select Customer Name");
        $("#productName").val("Select Product Name");
        $("#productType").val("");
        $("#productExpiryDate").val("");
        $("#productUnitPrice").val("");
        $("#noOfUnits").val("");
        $("#salesValue").val("");
    }
    catch(error){
      alert(error.message);
    }
  }

// Edit Item
  private async editItem() {
    try{
        var oId=0;
        var pId=0;
        var proId=0;
        var custId=0;
        var customerName='';
        var orderId = $("#orderId").val().toString();
        var orderIdFormat = /\bOID-\b\b[0-9]+\b/;
      
        var orderListUrl = sp.web.lists.getByTitle("Orders").items;
        var orderItems: any[] = await orderListUrl.getAll();
        orderItems.forEach((element) => {
          if (element.OrdersID == orderId) {
            oId = element.ID;
            proId = element.ProductID;
            console.log(proId);
            custId = element.CustomerID;
          }
        });
        if(orderId=='')
        alert("Please enter Order ID");
        else if(!orderIdFormat.test(orderId))
        alert("Enter Order Id in correct Format (OID-digits)");
        else if (oId == 0) 
        alert("Order Id not found");
        else if (proId == 0) 
        alert("Product id not found");
        else if (custId == 0) 
        alert("Customer Id not found");
        else
        {
          var ProductListUrl = sp.web.lists.getByTitle("Products").items;
          var products: any[] = await ProductListUrl.getAll();
          products.forEach((element) => {
            if (element.ProductID == proId) pId = element.ID;
          });
          var customers: any[] = await sp.web.lists
            .getByTitle("Customers")
            .items.getAll();

          var proName: any = $("#productName").val();
          var proType: any = $("#productType").val();
          var proExpdate: string = $("#productExpiryDate").val().toString();
          var unitPrice: any = $("#productUnitPrice").val();
          var noOfUnits: any = $("#noOfUnits").val();
          var salesValue: any = $("#salesValue").val();
          orderListUrl.getById(oId).get()
          .then((element) => {
              unitPrice = element.UnitPrice;
              noOfUnits = element.UnitSold;
              salesValue = element.SalesValue;

              $("#productUnitPrice").val(unitPrice);
              $("#noOfUnits").val(noOfUnits);
              $("#salesValue").val(salesValue);
            })
            .then(() => {
              ProductListUrl.getById(pId)
                .get()
                .then((element) => {
                  proName = element.ProductName;
                  proType = element.ProductType;
                  proExpdate = element.ProductExpiryDate;

                  $("#productExpiryDate").val(proExpdate.substr(0, 10));
                  $("#productName").val(proName);
                  $("#productType").val(proType);
                });
            })
            .then(() => {
              customers.forEach((element) => {
                if (element.CustomerID == custId) {
                  customerName = element.CustomerName;
                  $("#customerName").val(customerName);
                }
              });
            })
            .catch((error)=>alert(error.message)); 
            $('#update').css('display','inline');
          }    
        }
        catch(error){
          alert("Order Id Not found it may have been deleted"+error.message);
        }   
  }

  // Update Item
private async updateItem(){
  try{
      var custId=0;
      var proId=0;
      var oId=0;
      var orderId = $("#orderId").val().toString();
      var customers: any[] = await sp.web.lists.getByTitle("Customers").items.getAll();
      var ProductListUrl = sp.web.lists.getByTitle("Products").items;
      var products: any[] = await ProductListUrl.getAll();
      var orderListUrl = sp.web.lists.getByTitle("Orders").items;
      var orderItems: any[] = await orderListUrl.getAll();
        customers.forEach((element)=>{
          if(element.CustomerName==$('#customerName').val())
          custId=element.CustomerID;
        });
        products.forEach((element)=>{
          if(element.ProductName==$('#productName').val())
          proId=element.ProductID;
        });
        
        orderItems.forEach((element) => {
          if (element.OrdersID == orderId) {
            oId = element.ID;
            //  proId = element.ProductID;
            //  custId = element.CustomerID;
          }
        });
          orderListUrl.getById(oId).update({   
            CustomerID:custId,
            ProductID:proId,
            UnitSold:$('#noOfUnits').val(),
            UnitPrice:$('#productUnitPrice').val(),
            SalesValue:$('#salesValue').val(),
            OrderStatus:"Approved"
          })
          .then(()=>alert(`Item with orderId ${orderId} updated successfully`))
          .then(()=>{this.readOrderList();this.resetItems();$('#update').hide();})
          .catch((error)=>alert("Refresh to edit again"+'\n'+error.message));
    }
    catch(error){
      alert(error.message);
    }
}


  private async deleteItem() {
    try{
        var id=0;
        var orderId = $("#orderId").val().toString();
        var orderListUrl = sp.web.lists.getByTitle("Orders").items;
        var orderItems: any[] = await orderListUrl.getAll();
        var orderIdFormat = /\bOID-\b\b[0-9]+\b/;

        if (orderId == "") 
        alert("Please enter order Id");
        else if (orderIdFormat.test(orderId)) 
        {
          orderItems.forEach((element) => {
            if (element.OrdersID == orderId) 
            id = element.ID;
          });
          sp.web.lists.getByTitle("Orders").items.getById(id).recycle()
            .then(() => alert(`Item with orderId ${orderId} deleted successfully`))
            .then(() => this.readOrderList())
            .catch((e) => alert("Error While Deleting:-OrderID not Found"));
        } 
        else 
        {
          alert("Please enter order Id in Correct Format(OID-Digits)");
        }
    }
    catch(error){
      alert(error.message);
    }  
  } 
}
