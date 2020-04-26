using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace BBSPA_ListGenerator {

  class SharePointListFactory {

    #region "Variables to track ClientContext, Site and Web"

    static string siteUrl = ConfigurationManager.AppSettings["targetSiteUrl"];
    static ClientContext clientContext = new ClientContext(siteUrl);
    static Site siteCollection = clientContext.Site;
    static Web site = clientContext.Web;

    static SharePointListFactory() {
      string userName = ConfigurationManager.AppSettings["userName"];
      string password = ConfigurationManager.AppSettings["password"];
      SecureString securePassword = new SecureString();
      foreach (char c in password) {
        securePassword.AppendChar(c);
      };
      clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);
      clientContext.Load(site);
      clientContext.Load(site.Lists);
      clientContext.Load(site.ContentTypes);
      clientContext.ExecuteQuery();
    }

    #endregion

    #region "Variables and helper methods for site columns, content types and lists"

    static Field CreateSiteColumn(string fieldName, string fieldDisplayName, string fieldType) {

      Console.WriteLine("Creating " + fieldName + " site column...");

      // delete existing field if it exists
      try {
        Field fld = site.Fields.GetByInternalNameOrTitle(fieldName);
        fld.DeleteObject();
        clientContext.ExecuteQuery();
      }
      catch { }

      string fieldXML = @"<Field Name='" + fieldName + "' " +
                                "DisplayName='" + fieldDisplayName + "' " +
                                "Type='" + fieldType + "' " +
                                "Group='Critical Path Training' > " +
                         "</Field>";

      Field field = site.Fields.AddFieldAsXml(fieldXML, true, AddFieldOptions.DefaultValue);
      clientContext.Load(field);
      clientContext.ExecuteQuery();
      return field;
    }

    static void DeleteContentType(string contentTypeName) {

      try {
        foreach (var ct in site.ContentTypes) {
          if (ct.Name.Equals(contentTypeName)) {
            ct.DeleteObject();
            Console.WriteLine("Deleting existing " + ct.Name + " content type...");
            clientContext.ExecuteQuery();
            break;
          }
        }
      }
      catch { }

    }

    static ContentType CreateContentType(string contentTypeName, string baseContentType) {

      DeleteContentType(contentTypeName);

      ContentTypeCreationInformation contentTypeCreateInfo = new ContentTypeCreationInformation();
      contentTypeCreateInfo.Name = contentTypeName;
      contentTypeCreateInfo.ParentContentType = site.ContentTypes.GetById(baseContentType); ;
      contentTypeCreateInfo.Group = "Critical Path Training";
      ContentType ctype = site.ContentTypes.Add(contentTypeCreateInfo);
      clientContext.ExecuteQuery();
      return ctype;

    }

    static void DeleteList(string listTitle) {
      try {
        List list = site.Lists.GetByTitle(listTitle);
        list.DeleteObject();
        Console.WriteLine("Deleting existing " + listTitle + " list...");
        clientContext.ExecuteQuery();
      }
      catch { }
    }

    #endregion

    public static void CreateAllLists() {
      DeleteOrderDetailsList();
      DeleteOrdersList();
      DeleteCustomersList();
      CreateCustomersList(8, 8);
      CreateOrdersList();
      CreateOrderDetailsList();
      Console.WriteLine();
      Console.WriteLine("All lists have been created");
      Console.WriteLine();
    }

    #region "Customers List"

    static List listCustomers;

    public static void CreateCustomersList(int CustomerCount, int BatchSize) {

      Console.WriteLine("Creating customers list...");

      ListCreationInformation listInformationCustomers = new ListCreationInformation();
      listInformationCustomers.Title = "Customers";
      listInformationCustomers.Url = "Lists/Customers";
      listInformationCustomers.QuickLaunchOption = QuickLaunchOptions.On;
      listInformationCustomers.TemplateType = (int)ListTemplateType.GenericList;
      listCustomers = site.Lists.Add(listInformationCustomers);
      listCustomers.OnQuickLaunch = true;
      listCustomers.EnableAttachments = false;
      listCustomers.Update();
      clientContext.ExecuteQuery();

      var fldTitle = listCustomers.Fields.GetByInternalNameOrTitle("Title");
      fldTitle.Title = "Last Name";
      fldTitle.Update();
      clientContext.ExecuteQuery();

      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("FirstName"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("Company"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("EMail"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("WorkPhone"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("HomePhone"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("WorkAddress"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("WorkCity"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("WorkState"));
      listCustomers.Fields.Add(site.Fields.GetByInternalNameOrTitle("WorkZip"));
      clientContext.ExecuteQuery();

      listCustomers.DefaultView.ViewFields.Add("FirstName");
      listCustomers.DefaultView.ViewFields.Add("Company");
      listCustomers.DefaultView.ViewFields.Add("EMail");
      listCustomers.DefaultView.Update();
      clientContext.ExecuteQuery();

      clientContext.Load(listCustomers);
      clientContext.ExecuteQuery();

      PopulateCustomersList(CustomerCount, BatchSize);
    }

    static void DeleteCustomersList() {
      DeleteList("Customers");
    }

    static void PopulateCustomersList(int CustomerCount, int BatchSize) {
      Console.WriteLine("Adding sample Customers list items...");

      int customerCounter = 0;
      int batchCounter = 0;
      int batchStart = 1;

      var customers = RandomCustomerGenerator.GetCustomerList(CustomerCount);

      foreach (var customer in customers) {
        // increment counters
        customerCounter += 1;
        batchCounter += 1;
        // add new customer item
        ListItem newCustomer = listCustomers.AddItem(new ListItemCreationInformation());
        newCustomer["Title"] = customer.LastName;
        newCustomer["FirstName"] = customer.FirstName;
        newCustomer["Company"] = customer.Company;
        newCustomer["HomePhone"] = customer.HomePhone;
        newCustomer["WorkPhone"] = customer.WorkPhone;
        newCustomer["EMail"] = customer.EmailAddress;
        newCustomer["WorkAddress"] = customer.Address;
        newCustomer["WorkCity"] = customer.City;
        newCustomer["WorkState"] = customer.State;
        newCustomer["WorkZip"] = customer.ZipCode;
        newCustomer.Update();
        if (batchCounter >= BatchSize) {
          clientContext.ExecuteQuery();
          batchCounter = 0;
          batchStart = customerCounter + 1;
        }
      }
      clientContext.ExecuteQuery();

    }

    static bool CustomersListExists() {
      if (listCustomers == null || listCustomers.ServerObjectIsNull == null || listCustomers.ServerObjectIsNull == true) {
        try {
          listCustomers = clientContext.Web.Lists.GetByTitle("Customers");
          clientContext.Load(listCustomers);
          clientContext.ExecuteQuery();
          if (listCustomers == null || listCustomers.ServerObjectIsNull == null || listCustomers.ServerObjectIsNull == true) {
            return false;
          }
        }
        catch {
          return false;
        }
      }
      return true;
    }

    #endregion

    #region "Orders List"

    static List listOrders;

    public static void CreateOrdersList() {

      Console.WriteLine("Creating orders list...");

      if (CustomersListExists()) {
        ListCreationInformation listInformationOrders = new ListCreationInformation();
        listInformationOrders.Title = "Orders";
        listInformationOrders.Url = "Lists/Orders";
        listInformationOrders.QuickLaunchOption = QuickLaunchOptions.On;
        listInformationOrders.TemplateType = (int)ListTemplateType.GenericList;
        listOrders = site.Lists.Add(listInformationOrders);
        listOrders.OnQuickLaunch = true;
        listOrders.EnableAttachments = false;
        listOrders.Update();
        clientContext.ExecuteQuery();

        clientContext.Load(listOrders.DefaultView.ViewFields);
        clientContext.ExecuteQuery();

        listOrders.DefaultView.ViewFields.RemoveAll();
        listOrders.DefaultView.ViewFields.Add("ID");
        listOrders.DefaultView.ViewFields.Add("Title");
        listOrders.DefaultView.Update();
        clientContext.ExecuteQuery();

        string fldCustomerLookupXml = @"<Field Name='Customer' DisplayName='Customer' Type='Lookup' ></Field>";
        FieldLookup fldCustomerLookup =
          clientContext.CastTo<FieldLookup>(listOrders.Fields.AddFieldAsXml(fldCustomerLookupXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));

        // add cusotmer lookup field
        fldCustomerLookup.LookupField = "Title";
        fldCustomerLookup.LookupList = listCustomers.Id.ToString();
        fldCustomerLookup.Indexed = true;
        fldCustomerLookup.RelationshipDeleteBehavior = RelationshipDeleteBehaviorType.Cascade;
        fldCustomerLookup.Update();

        // add order date field
        string fldOrderDateXml = @"<Field Name='OrderDate' DisplayName='OrderDate' Type='DateTime' ></Field>";
        FieldDateTime fldOrderDate =
          clientContext.CastTo<FieldDateTime>(listOrders.Fields.AddFieldAsXml(fldOrderDateXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));
        fldOrderDate.DisplayFormat = DateTimeFieldFormatType.DateOnly;
        fldOrderDate.Update();

        // add order date field
        string fldOrderAmountXml = @"<Field Name='OrderAmount' DisplayName='OrderAmount' Type='Currency' ></Field>";
        FieldCurrency fldOrderAmount =
          clientContext.CastTo<FieldCurrency>(listOrders.Fields.AddFieldAsXml(fldOrderAmountXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));
        fldOrderAmount.Update();

        clientContext.ExecuteQuery();

        clientContext.Load(listOrders);
        clientContext.ExecuteQuery();

      }
      else {
        Console.WriteLine("Cannot create Orders list because Customer list does not exist.");
      }

    }

    public static void DeleteOrdersList() {
      if (OrdersListExists()) {
        // delete customer lookup column if it exists
        ExceptionHandlingScope scope = new ExceptionHandlingScope(clientContext);
        using (scope.StartScope()) {
          using (scope.StartTry()) {
            var col = listOrders.Fields.GetByInternalNameOrTitle("Customer");
            col.DeleteObject();
            listOrders.Update();
          }
          using (scope.StartCatch()) { }
        }
        clientContext.ExecuteQuery();

        // delete orders list
        DeleteList("Orders");
        listOrders = null;
      }
    }

    public static bool OrdersListExists() {
      if (listOrders == null || listOrders.ServerObjectIsNull == null || listOrders.ServerObjectIsNull == true) {
        try {
          listOrders = clientContext.Web.Lists.GetByTitle("Orders");
          clientContext.Load(listOrders);
          clientContext.ExecuteQuery();
          if (listOrders == null || (listOrders.ServerObjectIsNull == null && listOrders.ServerObjectIsNull == true)) {
            return false;
          }
        }
        catch {
          return false;
        }
      }
      return true;

    }

    #endregion

    #region "OrderDetails list"

    static List listOrderDetails;

    public static void CreateOrderDetailsList() {

      Console.WriteLine("Creating order details list...");

      if (OrdersListExists()) {
        ListCreationInformation listInformationOrderDetails = new ListCreationInformation();
        listInformationOrderDetails.Title = "OrderDetails";
        listInformationOrderDetails.Url = "Lists/OrderDetails";
        listInformationOrderDetails.QuickLaunchOption = QuickLaunchOptions.On;
        listInformationOrderDetails.TemplateType = (int)ListTemplateType.GenericList;
        listOrderDetails = site.Lists.Add(listInformationOrderDetails);
        listOrderDetails.OnQuickLaunch = true;
        listOrderDetails.EnableAttachments = false;
        listOrderDetails.Update();
        clientContext.ExecuteQuery();

        listOrderDetails.DefaultView.ViewFields.RemoveAll();
        listOrderDetails.DefaultView.ViewFields.Add("ID");
        listOrderDetails.DefaultView.Update();
        clientContext.ExecuteQuery();

        var fldTitle = listOrderDetails.Fields.GetByInternalNameOrTitle("Title");
        fldTitle.Required = false;
        fldTitle.Update();
        clientContext.ExecuteQuery();

        string fldOrderLookupXml = @"<Field Name='OrderId' DisplayName='Order' Type='Lookup' ></Field>";
        FieldLookup fldOrderLookup =
          clientContext.CastTo<FieldLookup>(listOrderDetails.Fields.AddFieldAsXml(fldOrderLookupXml,
                                                                            true,
                                                                            AddFieldOptions.AddFieldInternalNameHint));

        // add cusotmer lookup field
        fldOrderLookup.LookupField = "ID";
        fldOrderLookup.LookupList = listOrders.Id.ToString();
        fldOrderLookup.Indexed = true;
        fldOrderLookup.RelationshipDeleteBehavior = RelationshipDeleteBehaviorType.Cascade;
        fldOrderLookup.Update();

        // add quantity field
        string fldQuantityXml = @"<Field Name='Quantity' DisplayName='Quantity' Type='Number' ></Field>";
        FieldNumber fldQuantity =
          clientContext.CastTo<FieldNumber>(listOrderDetails.Fields.AddFieldAsXml(fldQuantityXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));
        fldQuantity.Update();

        // add product field
        string fldProductXml = @"<Field Name='Product' DisplayName='Product' Type='Text' ></Field>";
        FieldText fldProduct =
          clientContext.CastTo<FieldText>(listOrderDetails.Fields.AddFieldAsXml(fldProductXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));
        fldProduct.Update();

        string fldSalesAmountXml = @"<Field Name='SalesAmount' DisplayName='SalesAmount' Type='Currency' ></Field>";
        FieldCurrency fldSalesAmount =
          clientContext.CastTo<FieldCurrency>(listOrderDetails.Fields.AddFieldAsXml(fldSalesAmountXml,
                                                                            true,
                                                                            AddFieldOptions.DefaultValue));
        fldSalesAmount.Update();

        clientContext.ExecuteQuery();

        //listOrderDetails.DefaultView.ViewFields.Remove("Title");
        listOrderDetails.DefaultView.Update();
        clientContext.ExecuteQuery();
      }
      else {
        Console.WriteLine("Cannot create OrderDetails list because Orders list does not exist.");
      }

    }

    public static void DeleteOrderDetailsList() {
      if (OrderDetailsListExists()) {

        // delete customer lookup column if it exists
        ExceptionHandlingScope scope = new ExceptionHandlingScope(clientContext);
        using (scope.StartScope()) {
          using (scope.StartTry()) {
            var col = listOrderDetails.Fields.GetByInternalNameOrTitle("OrderId");
            col.DeleteObject();
            listOrderDetails.Update();
          }
          using (scope.StartCatch()) { }
        }

        DeleteList("OrderDetails");

      }
    }

    public static bool OrderDetailsListExists() {

      if (listOrderDetails == null || listOrderDetails.ServerObjectIsNull == null || listOrderDetails.ServerObjectIsNull == true) {
        try {
          listOrderDetails = clientContext.Web.Lists.GetByTitle("OrderDetails");
          clientContext.Load(listOrderDetails);
          clientContext.ExecuteQuery();
        }
        catch {
          return false;
        }
        if (listOrderDetails == null || (listOrderDetails.ServerObjectIsNull == null && listOrderDetails.ServerObjectIsNull == true)) {
          return false;
        }
      }
      return true;
    }

    #endregion

    #region "Expenses List"

    public static void CreateExpensesLists() {
      DeleteExpenseListTypes();
      CreateExpenseSiteColumns();
      CreateExpenseContentTypes();
      CreateExpensesList();
      CreateExpenseBudgetsList();
    }

    static FieldChoice fldExpenseCategory;
    static FieldDateTime fldExpenseDate;
    static FieldCurrency fldExpenseAmount;


    static FieldText fldExpenseBudgetYear;
    static FieldText fldExpenseBudgetQuarter;
    static FieldCurrency fldExpenseBudgetAmount;

    static ContentType ctypeExpense;
    static ContentType ctypeExpenseBudgetItem;

    static List listExpenses;
    static List listExpenseBudgets;

    static void DeleteExpenseListTypes() {
      DeleteList("Expenses");
      DeleteList("Expense Budgets");
      DeleteContentType("Expense Item");
      DeleteContentType("Expense Budget Item");
    }

    class ExpenseCategory {
      public const string OfficeSupplies = "Office Supplies";
      public const string Marketing = "Marketing";
      public const string Operations = "Operations";
      public const string ResearchAndDevelopment = "Research & Development";
      public static string[] GetAll() {
        string[] AllCategories = { OfficeSupplies, Marketing, Operations, ResearchAndDevelopment };
        return AllCategories;
      }
    }

    static void CreateExpenseSiteColumns() {

      fldExpenseCategory = clientContext.CastTo<FieldChoice>(CreateSiteColumn("ExpenseCategory", "Expense Category", "Choice"));
      string[] choicesExpenseCategory = ExpenseCategory.GetAll();
      fldExpenseCategory.Choices = choicesExpenseCategory;
      fldExpenseCategory.Update();
      clientContext.ExecuteQuery();


      fldExpenseDate = clientContext.CastTo<FieldDateTime>(CreateSiteColumn("ExpenseDate", "Expense Date", "DateTime")); ;
      fldExpenseDate.DisplayFormat = DateTimeFieldFormatType.DateOnly;
      fldExpenseDate.Update();

      fldExpenseAmount = clientContext.CastTo<FieldCurrency>(CreateSiteColumn("ExpenseAmount", "Expense Amount", "Currency"));
      fldExpenseAmount.MinimumValue = 0;

      fldExpenseBudgetYear = clientContext.CastTo<FieldText>(CreateSiteColumn("ExpenseBudgetYear", "Budget Year", "Text"));

      fldExpenseBudgetQuarter = clientContext.CastTo<FieldText>(CreateSiteColumn("ExpenseBudgetQuarter", "Budget Quarter", "Text"));
      fldExpenseBudgetQuarter.Update();

      fldExpenseBudgetAmount = clientContext.CastTo<FieldCurrency>(CreateSiteColumn("ExpenseBudgetAmount", "Budget Amount", "Currency"));

      clientContext.ExecuteQuery();
    }

    static void CreateExpenseContentTypes() {

      ctypeExpense = CreateContentType("Expense Item", "0x01");
      ctypeExpense.Update(true);
      clientContext.Load(ctypeExpense.FieldLinks);
      clientContext.ExecuteQuery();

      FieldLinkCreationInformation fldLinkExpenseCategory = new FieldLinkCreationInformation();
      fldLinkExpenseCategory.Field = fldExpenseCategory;
      ctypeExpense.FieldLinks.Add(fldLinkExpenseCategory);
      ctypeExpense.Update(true);

      // add site columns
      FieldLinkCreationInformation fldLinkExpenseDate = new FieldLinkCreationInformation();
      fldLinkExpenseDate.Field = fldExpenseDate;
      ctypeExpense.FieldLinks.Add(fldLinkExpenseDate);
      ctypeExpense.Update(true);

      // add site columns
      FieldLinkCreationInformation fldLinkExpenseAmount = new FieldLinkCreationInformation();
      fldLinkExpenseAmount.Field = fldExpenseAmount;
      ctypeExpense.FieldLinks.Add(fldLinkExpenseAmount);
      ctypeExpense.Update(true);

      clientContext.ExecuteQuery();

      ctypeExpenseBudgetItem = CreateContentType("Expense Budget Item", "0x01");
      ctypeExpenseBudgetItem.Update(true);
      clientContext.Load(ctypeExpenseBudgetItem.FieldLinks);
      clientContext.ExecuteQuery();

      FieldLinkCreationInformation fldLinkExpenseBudgetCategory = new FieldLinkCreationInformation();
      fldLinkExpenseBudgetCategory.Field = fldExpenseCategory;
      ctypeExpenseBudgetItem.FieldLinks.Add(fldLinkExpenseBudgetCategory);
      ctypeExpenseBudgetItem.Update(true);

      FieldLinkCreationInformation fldLinkExpenseBudgetYear = new FieldLinkCreationInformation();
      fldLinkExpenseBudgetYear.Field = fldExpenseBudgetYear;
      ctypeExpenseBudgetItem.FieldLinks.Add(fldLinkExpenseBudgetYear);
      ctypeExpenseBudgetItem.Update(true);

      FieldLinkCreationInformation fldLinkExpenseBudgetQuarter = new FieldLinkCreationInformation();
      fldLinkExpenseBudgetQuarter.Field = fldExpenseBudgetQuarter;
      ctypeExpenseBudgetItem.FieldLinks.Add(fldLinkExpenseBudgetQuarter);
      ctypeExpenseBudgetItem.Update(true);

      FieldLinkCreationInformation fldLinkExpenseBudgetAmount = new FieldLinkCreationInformation();
      fldLinkExpenseBudgetAmount.Field = fldExpenseBudgetAmount;
      ctypeExpenseBudgetItem.FieldLinks.Add(fldLinkExpenseBudgetAmount);
      ctypeExpenseBudgetItem.Update(true);

      clientContext.ExecuteQuery();

    }

    static void CreateExpensesList() {

      string listTitle = "Expenses";
      string listUrl = "Lists/Expenses";

      // delete document library if it already exists
      ExceptionHandlingScope scope = new ExceptionHandlingScope(clientContext);
      using (scope.StartScope()) {
        using (scope.StartTry()) {
          site.Lists.GetByTitle(listTitle).DeleteObject();
        }
        using (scope.StartCatch()) { }
      }

      ListCreationInformation lci = new ListCreationInformation();
      lci.Title = listTitle;
      lci.Url = listUrl;
      lci.TemplateType = (int)ListTemplateType.GenericList;
      listExpenses = site.Lists.Add(lci);
      listExpenses.OnQuickLaunch = true;
      listExpenses.EnableFolderCreation = false;
      listExpenses.Update();


      // attach JSLink script to default view for client-side rendering
      //listExpenses.DefaultView.JSLink = AppRootFolderRelativeUrl + "scripts/CustomersListCSR.js";
      listExpenses.DefaultView.Update();
      listExpenses.Update();
      clientContext.Load(listExpenses);
      clientContext.Load(listExpenses.Fields);
      var titleField = listExpenses.Fields.GetByInternalNameOrTitle("Title");
      titleField.Title = "Expense Description";
      titleField.Update();
      clientContext.ExecuteQuery();

      listExpenses.ContentTypesEnabled = true;
      listExpenses.ContentTypes.AddExistingContentType(ctypeExpense);
      listExpenses.Update();
      clientContext.Load(listExpenses.ContentTypes);
      clientContext.ExecuteQuery();

      ContentType existing = listExpenses.ContentTypes[0];
      existing.DeleteObject();
      clientContext.ExecuteQuery();

      View viewProducts = listExpenses.DefaultView;

      viewProducts.ViewFields.Add("ExpenseCategory");
      viewProducts.ViewFields.Add("ExpenseDate");
      viewProducts.ViewFields.Add("ExpenseAmount");
      viewProducts.Update();

      clientContext.ExecuteQuery();

      PopulateExpensesList();

    }

    static void CreateExpenseBudgetsList() {

      string listTitle = "Expense Budgets";
      string listUrl = "Lists/ExpenseBudgets";

      // delete document library if it already exists
      ExceptionHandlingScope scope = new ExceptionHandlingScope(clientContext);
      using (scope.StartScope()) {
        using (scope.StartTry()) {
          site.Lists.GetByTitle(listTitle).DeleteObject();
        }
        using (scope.StartCatch()) { }
      }

      ListCreationInformation lci = new ListCreationInformation();
      lci.Title = listTitle;
      lci.Url = listUrl;
      lci.TemplateType = (int)ListTemplateType.GenericList;
      listExpenseBudgets = site.Lists.Add(lci);
      listExpenseBudgets.OnQuickLaunch = true;
      listExpenseBudgets.EnableFolderCreation = false;
      listExpenseBudgets.Update();

      listExpenseBudgets.DefaultView.Update();
      listExpenseBudgets.Update();
      clientContext.Load(listExpenseBudgets);
      clientContext.Load(listExpenseBudgets.Fields);
      var titleField = listExpenseBudgets.Fields.GetByInternalNameOrTitle("Title");
      titleField.Title = "Expense Budget";
      titleField.Update();
      clientContext.ExecuteQuery();

      listExpenseBudgets.ContentTypesEnabled = true;
      listExpenseBudgets.ContentTypes.AddExistingContentType(ctypeExpenseBudgetItem);
      listExpenseBudgets.Update();
      clientContext.Load(listExpenseBudgets.ContentTypes);
      clientContext.ExecuteQuery();

      ContentType existing = listExpenseBudgets.ContentTypes[0];
      existing.DeleteObject();
      clientContext.ExecuteQuery();

      View viewProducts = listExpenseBudgets.DefaultView;

      viewProducts.ViewFields.Add("ExpenseCategory");
      viewProducts.ViewFields.Add("ExpenseBudgetYear");
      viewProducts.ViewFields.Add("ExpenseBudgetQuarter");
      viewProducts.ViewFields.Add("ExpenseBudgetAmount");
      viewProducts.Update();

      clientContext.ExecuteQuery();

      PopulateExpenseBudgetsList();

    }

    static void AddExpense(string Description, string Category, DateTime Date, decimal Amount) {

      ListItem newItem = listExpenses.AddItem(new ListItemCreationInformation());
      newItem["Title"] = Description;
      newItem["ExpenseCategory"] = Category;
      newItem["ExpenseDate"] = Date;
      newItem["ExpenseAmount"] = Amount;

      newItem.Update();
      clientContext.ExecuteQuery();

      Console.Write(".");
    }

    static void PopulateExpensesList() {

      Console.Write("Adding expenses");

      // January 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 1, 3), 133.44m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 1, 3), 328.40m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 1, 5), 824.90m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 1, 8), 89.40m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 1, 18), 23.90m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 1, 21), 478.33m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 1, 21), 20.00m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 1, 24), 12.50m);
      AddExpense("Toy Stress Tester", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 1, 28), 2400.00m);
      AddExpense("Office Depot supply run", ExpenseCategory.OfficeSupplies, new DateTime(2019, 1, 29), 184.30m);

      // Feb 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 2, 1), 138.02m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 2, 1), 297.47m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 2, 1), 789.77m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 8.95m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 74.55m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 45.67m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 32.34m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 20m);
      AddExpense("Toy Stress Tester", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 2, 1), 2400m);
      AddExpense("Office Depot supply run", ExpenseCategory.OfficeSupplies, new DateTime(2019, 2, 1), 196.44m);
      AddExpense("TV Ads - East Coast", ExpenseCategory.Marketing, new DateTime(2019, 2, 1), 2800m);
      AddExpense("TV Ads - West Coast", ExpenseCategory.Marketing, new DateTime(2019, 2, 1), 2400m);

      // March 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 3, 1), 142.99m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 3, 1), 304.21m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 3, 1), 804.33m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 44.23m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 3, 1), 500m);
      AddExpense("Printer Paper", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 48.20m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 20m);
      AddExpense("Toner Cartridges for Printer", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 220.34m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 8.95m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 3, 1), 12.30m);

      // April 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 4, 1), 138.34m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 4, 1), 344.32m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 4, 1), 812.90m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 4, 1), 32.45m);
      AddExpense("Toy Stress Tester", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 4, 1), 2400m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 4, 1), 500m);
      AddExpense("Print Ad in People Magazine", ExpenseCategory.Marketing, new DateTime(2019, 4, 1), 1200m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 4, 1), 34.20m);
      AddExpense("Toner Cartridges for Printer", ExpenseCategory.OfficeSupplies, new DateTime(2019, 4, 1), 127.88m);


      // May 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 5, 1), 152.55m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 5, 1), 320.45m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 5, 1), 783.44m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 5, 1), 23.90m);
      AddExpense("Toner Cartridges for Printer", ExpenseCategory.OfficeSupplies, new DateTime(2019, 5, 1), 240.50m);
      AddExpense("Printer Paper", ExpenseCategory.OfficeSupplies, new DateTime(2019, 5, 1), 22.32m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 5, 1), 20m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 5, 1), 8.95m);


      // June 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 6, 1), 138.44m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 6, 1), 332.78m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 6, 1), 802.44m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 6, 1), 34.22m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 6, 1), 8.95m);
      AddExpense("Print Ad in People Magazine", ExpenseCategory.Marketing, new DateTime(2019, 6, 1), 1200m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 6, 1), 24.10m);
      AddExpense("Toner Cartridges for Printer", ExpenseCategory.OfficeSupplies, new DateTime(2019, 6, 1), 132.20m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 6, 1), 8.95m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 6, 1), 500m);

      // July 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 7, 1), 135.22m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 7, 1), 333.11m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 7, 1), 798.25m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 7, 1), 8.95m);
      AddExpense("Office Depot supply run", ExpenseCategory.OfficeSupplies, new DateTime(2019, 7, 1), 212.41m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 7, 1), 46.78m);
      AddExpense("Particle Accelerator", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 7, 1), 4800m);

      // August 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 8, 1), 142.20m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 8, 1), 345.80m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 8, 1), 814.87m);
      AddExpense("TV Ads - Southeast", ExpenseCategory.Marketing, new DateTime(2019, 8, 1), 2800m);
      AddExpense("Toy Stress Tester", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 8, 1), 2400m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 8, 1), 8.95m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 8, 1), 500m);
      AddExpense("Server computer", ExpenseCategory.Operations, new DateTime(2019, 8, 1), 2500m);
      AddExpense("Office chairs", ExpenseCategory.OfficeSupplies, new DateTime(2019, 8, 1), 890.10m);


      // September 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 9, 1), 136.10m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 9, 1), 326.01m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 9, 1), 802.90m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 42.34m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 8.95m);
      AddExpense("Printer Paper", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 86.10m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 20m);
      AddExpense("Toner Cartridges for Printer", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 190.50m);
      AddExpense("Paper clips", ExpenseCategory.OfficeSupplies, new DateTime(2019, 9, 1), 8.95m);
      AddExpense("Server computer", ExpenseCategory.Operations, new DateTime(2019, 9, 1), 3200m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 9, 1), 500m);


      // October 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 10, 1), 141.33m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 10, 1), 322.55m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 10, 1), 832.50m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 10, 1), 35.34m);
      AddExpense("TV Ads - Southeast", ExpenseCategory.Marketing, new DateTime(2019, 10, 1), 4800m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 10, 1), 20m);
      AddExpense("Office Depot supply run", ExpenseCategory.OfficeSupplies, new DateTime(2019, 10, 1), 107.33m);
      AddExpense("Server computer", ExpenseCategory.Operations, new DateTime(2019, 10, 1), 2800m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 10, 1), 8.95m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 10, 1), 30.66m);
      AddExpense("Slide rule", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 10, 1), 48.50m);


      // November 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 11, 1), 140.10m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 11, 1), 321.98m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 11, 1), 842.90m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 11, 1), 42.11m);
      AddExpense("TV Ads - West Coast", ExpenseCategory.Marketing, new DateTime(2019, 11, 1), 4800m);
      AddExpense("File cabinet", ExpenseCategory.OfficeSupplies, new DateTime(2019, 11, 1), 120m);
      AddExpense("Printer Paper", ExpenseCategory.OfficeSupplies, new DateTime(2019, 11, 1), 220.34m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 11, 1), 500m);
      AddExpense("Postage Stamps", ExpenseCategory.OfficeSupplies, new DateTime(2019, 11, 1), 20m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 11, 1), 28.35m);

      // December 2019
      AddExpense("Water Bill", ExpenseCategory.Operations, new DateTime(2019, 12, 1), 326.48m);
      AddExpense("Verizon - Telephone Expenses", ExpenseCategory.Operations, new DateTime(2019, 12, 1), 345.32m);
      AddExpense("Electricity Bill", ExpenseCategory.Operations, new DateTime(2019, 12, 1), 840.66m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 8.95m);
      AddExpense("Printer Paper", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 34.20m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 12, 1), 500m);
      AddExpense("Cleaning Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 144.50m);
      AddExpense("Particle Accelerator", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 12, 1), 1200m);
      AddExpense("Pencils", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 8.95m);
      AddExpense("TV Ads - Southeast", ExpenseCategory.Marketing, new DateTime(2019, 12, 1), 1200m);
      AddExpense("Science Calculator", ExpenseCategory.ResearchAndDevelopment, new DateTime(2019, 12, 1), 120m);
      AddExpense("TV Ads - East Coast", ExpenseCategory.Marketing, new DateTime(2019, 12, 1), 1800m);
      AddExpense("TV Ads - West Coast", ExpenseCategory.Marketing, new DateTime(2019, 12, 1), 900m);
      AddExpense("Coffee Supplies", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 45.33m);
      AddExpense("Google Ad Words", ExpenseCategory.Marketing, new DateTime(2019, 12, 1), 500m);
      AddExpense("Office chairs", ExpenseCategory.OfficeSupplies, new DateTime(2019, 12, 1), 780.32m);

      Console.WriteLine();
      Console.WriteLine();
    }

    static void AddExpenseBudget(string Category, string Year, string Quarter, decimal Amount) {

      ListItem newItem = listExpenseBudgets.AddItem(new ListItemCreationInformation());
      newItem["Title"] = Category + " for " + Quarter.ToString() + " of " + Year.ToString();
      newItem["ExpenseCategory"] = Category;
      newItem["ExpenseBudgetYear"] = Year;
      newItem["ExpenseBudgetQuarter"] = Quarter;
      newItem["ExpenseBudgetAmount"] = Amount;

      newItem.Update();
      clientContext.ExecuteQuery();

      Console.Write(".");
    }

    static void PopulateExpenseBudgetsList() {

      Console.Write("Adding expense budgets");

      AddExpenseBudget(ExpenseCategory.OfficeSupplies, "2019", "Q1", 1000m);
      AddExpenseBudget(ExpenseCategory.Marketing, "2019", "Q1", 7500m);
      AddExpenseBudget(ExpenseCategory.Operations, "2019", "Q1", 7000m);
      AddExpenseBudget(ExpenseCategory.ResearchAndDevelopment, "2019", "Q1", 5000m);

      AddExpenseBudget(ExpenseCategory.OfficeSupplies, "2019", "Q2", 1000m);
      AddExpenseBudget(ExpenseCategory.Marketing, "2019", "Q2", 7500m);
      AddExpenseBudget(ExpenseCategory.Operations, "2019", "Q2", 7000m);
      AddExpenseBudget(ExpenseCategory.ResearchAndDevelopment, "2019", "Q2", 5000m);

      AddExpenseBudget(ExpenseCategory.OfficeSupplies, "2019", "Q3", 1000m);
      AddExpenseBudget(ExpenseCategory.Marketing, "2019", "Q3", 10000m);
      AddExpenseBudget(ExpenseCategory.Operations, "2019", "Q3", 7000m);
      AddExpenseBudget(ExpenseCategory.ResearchAndDevelopment, "2019", "Q3", 5000m);

      AddExpenseBudget(ExpenseCategory.OfficeSupplies, "2019", "Q4", 1000m);
      AddExpenseBudget(ExpenseCategory.Marketing, "2019", "Q4", 10000m);
      AddExpenseBudget(ExpenseCategory.Operations, "2019", "Q4", 7000m);
      AddExpenseBudget(ExpenseCategory.ResearchAndDevelopment, "2019", "Q4", 5000m);


      Console.WriteLine();
      Console.WriteLine();
    }

    #endregion

  }
}
