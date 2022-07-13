var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from "react";
import styles from "./MfrpSalesForm.module.scss";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as $ from "jquery";
var MfrpSalesForm = /** @class */ (function (_super) {
    __extends(MfrpSalesForm, _super);
    //Constructor
    function MfrpSalesForm(props, state) {
        var _this = _super.call(this, props) || this;
        sp.setup({
            spfxContext: _this.props.spcontext,
        });
        state = {
            items: [],
            Products: [],
        };
        return _this;
    }
    MfrpSalesForm.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.mfrpSalesForm },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("h2", { className: styles.title }, "Sales Order Form"),
                    React.createElement("div", null, "===========================================================================================")),
                React.createElement("div", { className: styles.row },
                    React.createElement("label", { className: styles.label, htmlFor: "customerName" }, "Customer Name"),
                    React.createElement("select", { className: styles.input, id: "customerName", placeholder: "Customer Name" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "productName" }, "Product Name"),
                    React.createElement("select", { className: styles.input, id: "productName", placeholder: "Product Name" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "productType" }, "Product Type"),
                    React.createElement("input", { className: styles.input, id: "productType" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "productExpiryDate" }, "Product Expiry date"),
                    React.createElement("input", { className: styles.input, id: "productExpiryDate" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "productUnitPrice" }, "Product Unit Price"),
                    React.createElement("input", { className: styles.input, id: "productUnitPrice" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "noOfUnits" }, "Number of Units"),
                    React.createElement("input", { className: styles.input, id: "noOfUnits" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "salesValue" }, "Sales Value"),
                    React.createElement("input", { className: styles.input, id: "salesValue" }),
                    React.createElement("br", null),
                    React.createElement("br", null),
                    React.createElement("button", { className: styles.add, id: "add", onClick: function () { return _this.addItems(); } }, "Add"),
                    React.createElement("button", { className: styles.update, id: "update", onClick: function () { return _this.updateItem(); } }, "Update"),
                    React.createElement("button", { className: styles.reset, id: "reset", onClick: function () { return _this.resetItems(); } }, "Reset"),
                    React.createElement("br", null),
                    React.createElement("h2", { className: styles.title }, "Edit OR Delete Order"),
                    React.createElement("div", null, "==========================================================================================="),
                    React.createElement("br", null),
                    React.createElement("label", { className: styles.label, htmlFor: "orderId" }, "Enter Order Id"),
                    React.createElement("input", { className: styles.input, id: "orderId", placeholder: "Format=OID-digits" }),
                    React.createElement("button", { className: styles.edit, id: "edit", onClick: function () { return _this.editItem(); } }, "Edit"),
                    React.createElement("button", { className: styles.delete, id: "delete", onClick: function () { return _this.deleteItem(); } }, "Delete"),
                    React.createElement("h2", { className: styles.title }, "Orders List Items"),
                    React.createElement("div", { className: styles.row },
                        React.createElement("table", { id: "orderList" }))))));
    };
    MfrpSalesForm.prototype.componentDidMount = function () {
        this.readOrderList();
        this.Dropdowns();
    };
    // Read Order List Items
    MfrpSalesForm.prototype.readOrderList = function () {
        return __awaiter(this, void 0, void 0, function () {
            var orderItems, orderHtml, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists.getByTitle("Orders").items.getAll()];
                    case 1:
                        orderItems = _a.sent();
                        orderHtml = "<tr><th>Orders Id</th>\n                  <th>Customer Id</th>  \n                  <th>Product Id</th>\n                  <th>Unit Sold</th>   \n                  <th>Unit Price</th>   \n                  <th>Sales Value</th>\n                  <th>Order Status</th>\n                  </tr>";
                        orderItems.forEach(function (element) {
                            orderHtml += "<tr><td>" + element.OrdersID + "</td><td>" + element.CustomerID + "</td>\n                  <td>" + element.ProductID + "</td> <td>" + element.UnitPrice + "</td>\n                  <td>" + element.UnitSold + "</td><td>" + element.SalesValue + "</td>\n                  <td>" + element.OrderStatus + "</td></tr>";
                        });
                        $("#orderList").html(orderHtml);
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        alert(error_1.message);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    // DropDowns
    MfrpSalesForm.prototype.Dropdowns = function () {
        return __awaiter(this, void 0, void 0, function () {
            var items, customerItems, html, products, productItems, prohtml, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        items = [];
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Customers")
                                .items.getAll()];
                    case 1:
                        customerItems = _a.sent();
                        customerItems.forEach(function (element) {
                            items.push({ customerName: element.CustomerName });
                        });
                        this.setState({ items: items });
                        html = "<option>Select Customer Name</option>";
                        items.forEach(function (item) {
                            html += "<option>" + item.customerName + "</option>";
                        });
                        $("#customerName").append(html);
                        products = [];
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Products")
                                .items.getAll()];
                    case 2:
                        productItems = _a.sent();
                        this.setState({ Products: productItems });
                        productItems.forEach(function (element) {
                            products.push({ productName: element.ProductName });
                        });
                        prohtml = "<option>Select Product Name</option>";
                        products.forEach(function (item) {
                            prohtml += "<option>" + item.productName + "</option>";
                        });
                        $("#productName").append(prohtml);
                        this.AutoPopulate();
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _a.sent();
                        alert(error_2.message);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    //Auto populate
    MfrpSalesForm.prototype.AutoPopulate = function () {
        return __awaiter(this, void 0, void 0, function () {
            var productItems, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.lists.getByTitle("Products").items.getAll()];
                    case 1:
                        productItems = _a.sent();
                        $("#productName").on("change", function () {
                            var proName = $("#productName option:selected").val();
                            productItems.forEach(function (element) {
                                if (element.ProductName == proName) {
                                    // alert("found")
                                    var date = element.ProductExpiryDate;
                                    $("#productType").val(element.ProductType);
                                    $("#productUnitPrice").val(element.ProductUnitPrice);
                                    $("#productExpiryDate").val(date.substr(0, 10));
                                }
                                //  else
                                //  alert("Product Not found")
                            });
                            if ($('#noOfUnits').val() != "") {
                                var unitPrice = parseFloat($('#productUnitPrice').val().toString());
                                var units = parseFloat($('#noOfUnits').val().toString());
                                var sales = unitPrice * units;
                                $('#salesValue').val(sales);
                            }
                        });
                        $("#noOfUnits").on("change", function () {
                            var units = parseFloat($("#noOfUnits").val().toString());
                            var price = parseFloat($("#productUnitPrice").val().toString());
                            if ($("#noOfUnits").val() == "") {
                                alert("Please enter no of units");
                            }
                            else {
                                var sales = units * price;
                                $("#salesValue").val(sales);
                            }
                        });
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        alert(error_3.message);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    //  Add items to order list
    MfrpSalesForm.prototype.addItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var title, orderId, custId, proId, unitSold, unitPrice, salesValue, id, custName, proName, orderListUrl, orderItems, products, customers, titleNum, error_4;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        title = '';
                        orderId = '';
                        custId = '';
                        proId = '';
                        unitSold = $("#noOfUnits").val().toString();
                        unitPrice = $("#productUnitPrice").val();
                        salesValue = $("#salesValue").val();
                        id = 0;
                        custName = $("#customerName").val();
                        proName = $("#productName").val();
                        if (!(custName == "Select Customer Name")) return [3 /*break*/, 1];
                        alert("Please Select Customer Name");
                        return [3 /*break*/, 7];
                    case 1:
                        if (!(proName == "Select Product Name")) return [3 /*break*/, 2];
                        alert("Please Select Product Name");
                        return [3 /*break*/, 7];
                    case 2:
                        if (!(unitSold == "")) return [3 /*break*/, 3];
                        alert("please input number of units sold");
                        return [3 /*break*/, 7];
                    case 3:
                        orderListUrl = sp.web.lists.getByTitle("Orders").items;
                        return [4 /*yield*/, orderListUrl.getAll()];
                    case 4:
                        orderItems = _a.sent();
                        return [4 /*yield*/, sp.web.lists.getByTitle("Products").items.getAll()];
                    case 5:
                        products = _a.sent();
                        return [4 /*yield*/, sp.web.lists.getByTitle("Customers").items.getAll()];
                    case 6:
                        customers = _a.sent();
                        if (orderItems.length == 0) {
                            id = 0;
                            title = "1";
                        }
                        else {
                            orderItems.forEach(function (element) {
                                id = element.ID;
                                title = element.Title;
                            });
                        }
                        customers.forEach(function (element) {
                            if (element.CustomerName == custName) {
                                custId = element.CustomerID;
                            }
                        });
                        products.forEach(function (element) {
                            if (element.ProductName == proName) {
                                proId = element.ProductID;
                            }
                        });
                        titleNum = parseInt(title);
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
                            .then(function () { return alert("New Item with orderId " + orderId + " added successfully to Order List"); })
                            .then(function () { return _this.readOrderList(); })
                            .catch(function (e) { return alert("Error" + e.message); });
                        this.resetItems();
                        _a.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        error_4 = _a.sent();
                        alert(error_4.message);
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    // Reset Items in input Fields
    MfrpSalesForm.prototype.resetItems = function () {
        try {
            $("#customerName").val("Select Customer Name");
            $("#productName").val("Select Product Name");
            $("#productType").val("");
            $("#productExpiryDate").val("");
            $("#productUnitPrice").val("");
            $("#noOfUnits").val("");
            $("#salesValue").val("");
        }
        catch (error) {
            alert(error.message);
        }
    };
    // Edit Item
    MfrpSalesForm.prototype.editItem = function () {
        return __awaiter(this, void 0, void 0, function () {
            var oId, pId, proId, custId, customerName, orderId, orderIdFormat, orderListUrl, orderItems, ProductListUrl, products, customers, proName, proType, proExpdate, unitPrice, noOfUnits, salesValue, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 10, , 11]);
                        oId = 0;
                        pId = 0;
                        proId = 0;
                        custId = 0;
                        customerName = '';
                        orderId = $("#orderId").val().toString();
                        orderIdFormat = /\bOID-\b\b[0-9]+\b/;
                        orderListUrl = sp.web.lists.getByTitle("Orders").items;
                        return [4 /*yield*/, orderListUrl.getAll()];
                    case 1:
                        orderItems = _a.sent();
                        orderItems.forEach(function (element) {
                            if (element.OrdersID == orderId) {
                                oId = element.ID;
                                proId = element.ProductID;
                                console.log(proId);
                                custId = element.CustomerID;
                            }
                        });
                        if (!(orderId == '')) return [3 /*break*/, 2];
                        alert("Please enter Order ID");
                        return [3 /*break*/, 9];
                    case 2:
                        if (!!orderIdFormat.test(orderId)) return [3 /*break*/, 3];
                        alert("Enter Order Id in correct Format (OID-digits)");
                        return [3 /*break*/, 9];
                    case 3:
                        if (!(oId == 0)) return [3 /*break*/, 4];
                        alert("Order Id not found");
                        return [3 /*break*/, 9];
                    case 4:
                        if (!(proId == 0)) return [3 /*break*/, 5];
                        alert("Product id not found");
                        return [3 /*break*/, 9];
                    case 5:
                        if (!(custId == 0)) return [3 /*break*/, 6];
                        alert("Customer Id not found");
                        return [3 /*break*/, 9];
                    case 6:
                        ProductListUrl = sp.web.lists.getByTitle("Products").items;
                        return [4 /*yield*/, ProductListUrl.getAll()];
                    case 7:
                        products = _a.sent();
                        products.forEach(function (element) {
                            if (element.ProductID == proId)
                                pId = element.ID;
                        });
                        return [4 /*yield*/, sp.web.lists
                                .getByTitle("Customers")
                                .items.getAll()];
                    case 8:
                        customers = _a.sent();
                        proName = $("#productName").val();
                        proType = $("#productType").val();
                        proExpdate = $("#productExpiryDate").val().toString();
                        unitPrice = $("#productUnitPrice").val();
                        noOfUnits = $("#noOfUnits").val();
                        salesValue = $("#salesValue").val();
                        orderListUrl.getById(oId).get()
                            .then(function (element) {
                            unitPrice = element.UnitPrice;
                            noOfUnits = element.UnitSold;
                            salesValue = element.SalesValue;
                            $("#productUnitPrice").val(unitPrice);
                            $("#noOfUnits").val(noOfUnits);
                            $("#salesValue").val(salesValue);
                        })
                            .then(function () {
                            ProductListUrl.getById(pId)
                                .get()
                                .then(function (element) {
                                proName = element.ProductName;
                                proType = element.ProductType;
                                proExpdate = element.ProductExpiryDate;
                                $("#productExpiryDate").val(proExpdate.substr(0, 10));
                                $("#productName").val(proName);
                                $("#productType").val(proType);
                            });
                        })
                            .then(function () {
                            customers.forEach(function (element) {
                                if (element.CustomerID == custId) {
                                    customerName = element.CustomerName;
                                    $("#customerName").val(customerName);
                                }
                            });
                        })
                            .catch(function (error) { return alert(error.message); });
                        $('#update').css('display', 'inline');
                        _a.label = 9;
                    case 9: return [3 /*break*/, 11];
                    case 10:
                        error_5 = _a.sent();
                        alert("Order Id Not found it may have been deleted" + error_5.message);
                        return [3 /*break*/, 11];
                    case 11: return [2 /*return*/];
                }
            });
        });
    };
    // Update Item
    MfrpSalesForm.prototype.updateItem = function () {
        return __awaiter(this, void 0, void 0, function () {
            var custId, proId, oId, orderId, customers, ProductListUrl, products, orderListUrl, orderItems, error_6;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        custId = 0;
                        proId = 0;
                        oId = 0;
                        orderId = $("#orderId").val().toString();
                        return [4 /*yield*/, sp.web.lists.getByTitle("Customers").items.getAll()];
                    case 1:
                        customers = _a.sent();
                        ProductListUrl = sp.web.lists.getByTitle("Products").items;
                        return [4 /*yield*/, ProductListUrl.getAll()];
                    case 2:
                        products = _a.sent();
                        orderListUrl = sp.web.lists.getByTitle("Orders").items;
                        return [4 /*yield*/, orderListUrl.getAll()];
                    case 3:
                        orderItems = _a.sent();
                        customers.forEach(function (element) {
                            if (element.CustomerName == $('#customerName').val())
                                custId = element.CustomerID;
                        });
                        products.forEach(function (element) {
                            if (element.ProductName == $('#productName').val())
                                proId = element.ProductID;
                        });
                        orderItems.forEach(function (element) {
                            if (element.OrdersID == orderId) {
                                oId = element.ID;
                                //  proId = element.ProductID;
                                //  custId = element.CustomerID;
                            }
                        });
                        orderListUrl.getById(oId).update({
                            CustomerID: custId,
                            ProductID: proId,
                            UnitSold: $('#noOfUnits').val(),
                            UnitPrice: $('#productUnitPrice').val(),
                            SalesValue: $('#salesValue').val(),
                            OrderStatus: "Approved"
                        })
                            .then(function () { return alert("Item with orderId " + orderId + " updated successfully"); })
                            .then(function () { _this.readOrderList(); _this.resetItems(); $('#update').hide(); })
                            .catch(function (error) { return alert("Refresh to edit again" + '\n' + error.message); });
                        return [3 /*break*/, 5];
                    case 4:
                        error_6 = _a.sent();
                        alert(error_6.message);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    MfrpSalesForm.prototype.deleteItem = function () {
        return __awaiter(this, void 0, void 0, function () {
            var id, orderId, orderListUrl, orderItems, orderIdFormat, error_7;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        id = 0;
                        orderId = $("#orderId").val().toString();
                        orderListUrl = sp.web.lists.getByTitle("Orders").items;
                        return [4 /*yield*/, orderListUrl.getAll()];
                    case 1:
                        orderItems = _a.sent();
                        orderIdFormat = /\bOID-\b\b[0-9]+\b/;
                        if (orderId == "")
                            alert("Please enter order Id");
                        else if (orderIdFormat.test(orderId)) {
                            orderItems.forEach(function (element) {
                                if (element.OrdersID == orderId)
                                    id = element.ID;
                            });
                            sp.web.lists.getByTitle("Orders").items.getById(id).recycle()
                                .then(function () { return alert("Item with orderId " + orderId + " deleted successfully"); })
                                .then(function () { return _this.readOrderList(); })
                                .catch(function (e) { return alert("Error While Deleting:-OrderID not Found"); });
                        }
                        else {
                            alert("Please enter order Id in Correct Format(OID-Digits)");
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        error_7 = _a.sent();
                        alert(error_7.message);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return MfrpSalesForm;
}(React.Component));
export default MfrpSalesForm;
//# sourceMappingURL=MfrpSalesForm.js.map