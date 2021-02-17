using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace _18600351
{
    public partial class NorthwindForm : Form
    {
        DbProviderFactory factory;
        DbConnection connection;

        NorthwindDataContext northwind = new NorthwindDataContext();

        DataSet dataSet = new DataSet();

        BindingSource bindingSourceCategories = new BindingSource();
        BindingSource bindingSourceProducts = new BindingSource();
        BindingSource bindingSourceCustomers = new BindingSource();
        BindingSource bindingSourceOrders = new BindingSource();
        BindingSource bindingSourceOrderDetails = new BindingSource();
        BindingSource bindingSourceOrderedProducts = new BindingSource();

        string tableCategoriesName = "Categories";
        string tableProductsName = "Products";
        string tableCategories_Products_RelationnName = "Categories_Products_Relation";

        public class PagingInfo
        {
            public int RowsPerPage { get; set; }
            public int CurrentPage { get; set; }
            public int TotalPages { get; set; }
            public int TotalItems { get; set; }
            public List<string> Pages
            {
                get
                {
                    var result = new List<string>();

                    for (var i = 1; i <= TotalPages; i++)
                    {
                        result.Add($"Trang {i} / {TotalPages}");
                    }
                    return result;
                }
            }
        }

        PagingInfo pagingInfoProducts;
        PagingInfo pagingInfoCustomers;

        //Presentation
        public NorthwindForm()
        {
            InitializeComponent();
        }

        private void NorthwindForm_Loaded(object sender, EventArgs e)
        {
            ConnectionStringSettingsCollection settings = ConfigurationManager.ConnectionStrings;
            var connectionStrings = new List<ConnectionStringSettings>();

            if (settings != null)
            {
                foreach (ConnectionStringSettings cs in settings)
                {
                    if (!cs.Name.Contains("SQL_"))
                        continue;
                    connectionStrings.Add(cs);
                }
                databaseServersCombobox.DataSource = connectionStrings;
                databaseServersCombobox.DisplayMember = "name";
            }
        }

        private void loadButton_Click(object sender, EventArgs e)
        {
            var css = databaseServersCombobox.SelectedItem as ConnectionStringSettings;        
            try { loadDatabase(css); }
            catch (Exception) { Application.Restart(); }
        }

        private void addCategoryButton_Click(object sender, EventArgs e)
        {
            try
            {
                string categoryName = categoryNameTextBox.Text;
                string description = categoryDesscriptionTextBox.Text;
                byte[] picture = ImageToByteArray(picturePictureBox.Image, picturePictureBox);

                addCategory(categoryName, description, picture);

                dataSet.Tables[tableCategoriesName].Clear();
                loadTable(tableCategoriesName, tableCategoriesName);
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void updateCategoryButton_Click(object sender, EventArgs e)
        {
            try
            {
                int categoryId = int.Parse(categoriesCombobox.SelectedValue.ToString());
                string categoryName = categoryNameTextBox.Text;
                string description = categoryDesscriptionTextBox.Text;
                byte[] picture = ImageToByteArray(picturePictureBox.Image, picturePictureBox);

                updateCategory(categoryId, categoryName, description, picture);

                dataSet.Tables[tableCategoriesName].Clear();
                loadTable(tableCategoriesName, tableCategoriesName);
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void deleteCategoryButton_Click(object sender, EventArgs e)
        {
            try
            {
                int categoryId = int.Parse(categoriesCombobox.SelectedValue.ToString());
                deleteCategory(categoryId);

                dataSet.Tables[tableCategoriesName].Clear();
                loadTable(tableCategoriesName, tableCategoriesName);
                calculatePagingProducts();
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void browsePictureButton_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Choose Image";
            openFileDialog.Filter = "Images (*.JPEG;*.BMP;*.JPG;*.GIF;*.PNG;*.)|*.JPEG;*.BMP;*.JPG;*.GIF;*.PNG";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Image img = new Bitmap(openFileDialog.FileName);
                picturePictureBox.Image = (Image)new Bitmap(img, new Size(picturePictureBox.Width, picturePictureBox.Height));
            }
        }

        private void addProductButton_Click(object sender, EventArgs e)
        {
            try
            {
                string productName = productNameTextBox.Text;
                int categoryId = int.Parse(categoryCombobox.SelectedValue.ToString());
                string quantityPerUnit = quantityPerUnitTextBox.Text;
                decimal unitPrice = decimal.Parse(unitPriceTextBox.Text.ToString());
                short unitsInStock = short.Parse(UnitsInStockTextBox.Text.ToString());
                short unitsOnOrder = short.Parse(UnitsOnOrderTextBox.Text.ToString());
                short reorderLevel = short.Parse(ReorderLevelTextBox.Text.ToString());
                bool discontinued = discontinuedCheckBox.Checked;

                addProducts(productName, categoryId, quantityPerUnit, unitPrice, unitsInStock, unitsOnOrder, reorderLevel, discontinued);

                calculatePagingProducts();
                updateProductView();

                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void updateProductButton_Click(object sender, EventArgs e)
        {
            try
            {
                int productId = int.Parse(productIdTextBox.Text.ToString());
                string productName = productNameTextBox.Text;
                int categoryId = int.Parse(categoryCombobox.SelectedValue.ToString());
                string quantityPerUnit = quantityPerUnitTextBox.Text;
                decimal unitPrice = decimal.Parse(unitPriceTextBox.Text.ToString());
                short unitsInStock = short.Parse(UnitsInStockTextBox.Text.ToString());
                short unitsOnOrder = short.Parse(UnitsOnOrderTextBox.Text.ToString());
                short reorderLevel = short.Parse(ReorderLevelTextBox.Text.ToString());
                bool discontinued = discontinuedCheckBox.Checked;

                updateProducts(productId, productName, categoryId, quantityPerUnit, unitPrice, unitsInStock, unitsOnOrder, reorderLevel, discontinued);

                calculatePagingProducts();
                updateProductView();

                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void deleteProductButton_Click(object sender, EventArgs e)
        {
            try
            {
                int productId = int.Parse(productIdTextBox.Text.ToString());
                deleteProduct(productId);

                calculatePagingProducts();
                updateProductView();

                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void addCustomerButton_Click(object sender, EventArgs e)
        {
            try
            {
                string companyName = companyNameTextBox.Text;
                string contactName = contactNameTextBox.Text;
                string contactTitle = contactTitleTextBox.Text;
                string address = addressTextBox.Text;
                string city = cityTextBox.Text;
                string region = regionTextBox.Text;
                string postalCode = postalCodeTextBox.Text;
                string country = countryTextBox.Text;
                string phone = phoneTextBox.Text;
                string fax = faxTextBox.Text;

                addCustomer(companyName, contactName, contactTitle, address, city, region, postalCode, country, phone, fax);

                calculatePagingCustomers();
                updateProductView();
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void updateCustomerButton_Click(object sender, EventArgs e)
        {
            try
            {
                string customerId = customerIdTextBox.Text;
                string companyName = companyNameTextBox.Text;
                string contactName = contactNameTextBox.Text;
                string contactTitle = contactTitleTextBox.Text;
                string address = addressTextBox.Text;
                string city = cityTextBox.Text;
                string region = regionTextBox.Text;
                string postalCode = postalCodeTextBox.Text;
                string country = countryTextBox.Text;
                string phone = phoneTextBox.Text;
                string fax = faxTextBox.Text;

                updateCustomer(customerId, companyName, contactName, contactTitle, address, city, region, postalCode, country, phone, fax);

                calculatePagingCustomers();
                updateCustomersView();
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void deleteCustomerButton_Click(object sender, EventArgs e)
        {
            try
            {
                string customerId = customerIdTextBox.Text;
                deleteCustomer(customerId);

                calculatePagingCustomers();
                updateCustomersView();
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void seeOrdersButton_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(OrdersTabPage);
        }

        private void sortOrderByCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var shipCity = filerByShipCityCombobox.SelectedIndex == filerByShipCityCombobox.Items.Count - 1 ? "" : filerByShipCityCombobox.SelectedValue.ToString();

            var orders = northwind.Customers.ToList()[int.Parse(customerIdCombobox.SelectedIndex.ToString())].Orders;
            var query = from order in orders
                        where order.OrderDate >= filterFromOrderDateDateTimePicker.Value
                        && order.OrderDate <= filterToOrderDateDateTimePicker.Value
                        && order.ShipCity.ToLower().Contains(shipCity.ToLower())
                        select order;

            switch (sortOrdersByCombobox.SelectedIndex)
            {
                case 0:
                    //ordersDataGridView.Sort(ordersDataGridView.Columns["OrderDate"], ListSortDirection.Ascending);
                    bindingSourceOrders.DataSource = query.ToList().OrderBy(o => o.OrderDate);
                    break;
                case 1:
                    //ordersDataGridView.Sort(ordersDataGridView.Columns["OrderDate"], ListSortDirection.Descending);
                    bindingSourceOrders.DataSource = query.ToList().OrderByDescending(o => o.OrderDate);
                    break;
            }
        }

        private void createOrderButton_Click(object sender, EventArgs e)
        {
            try
            {
                int orderId = int.Parse(orderIdTextBox.Text);
                string customerId = orderCustomerIdTextBox.Text;
                DateTime orderDate = orderDateDateTimePicker.Value;
                DateTime requireDate = requireDateDateTimePicker.Value;
                DateTime shippedDate = shippedDateDateTimePicker.Value;
                decimal freight = decimal.Parse(freightTextBox.Text);
                string shipName = shipNameTextBox.Text;
                string shipAddress = addressTextBox.Text;
                string shipCity = shipCityTextBox.Text;
                string shipRegion = shipRegionTextBox.Text;
                string shipPostalCode = postalCodeTextBox.Text;
                string shipCountry = shipCountryTextBox.Text;
                addOrders(orderId, customerId, orderDate, requireDate, shippedDate, freight, shipName, shipAddress, shipCity, shipRegion, shipPostalCode, shipCountry);
                createOrderButton.Enabled = false;
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void createOrderDetailButton_Click(object sender, EventArgs e)
        {
            try
            {
                int orderId = int.Parse(orderIdCombobox.SelectedValue.ToString());
                int productId = int.Parse(orderProductsCombobox.SelectedValue.ToString());
                decimal unitPrice = decimal.Parse(orderUnitPriceTextBox.Text);
                short quantity = short.Parse(quantityTextBox.Text);
                float discount = float.Parse(discountTextBox.Text);
                addOrderDetail(orderId, productId, unitPrice, quantity, discount);
                createOrderDetailButton.Enabled = false;

                loadRevenueStatisticsByCustomers();
                loadRevenueStatisticsByProduct();
                loadRevenueStatisticsOverMonth();
                loadRevenueStatisticsOverYear();
                loadStatisticsOfProductsSoldOverMonthChart();
                loadStatisticsOfProductsSoldOverYearChart();
                MessageBox.Show("--successfully--");
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
                return;
            }
        }

        private void resetOrderButton_Click(object sender, EventArgs e)
        {
            orderIdTextBox.Text = autoGenerateIdOrder().ToString();
            freightTextBox.Text = "";
            shipNameTextBox.Text = "";
            shipAddressTextBox.Text = "";
            shipAddressTextBox.Text = "";
            shipCityTextBox.Text = "";
            shipRegionTextBox.Text = "";
            shipRegionTextBox.Text = "";
            shipPostalCodeTextBox.Text = "";
            shipCountryTextBox.Text = "";
            createOrderButton.Enabled = true;
        }

        private void resetOrderDetailsButton_Click(object sender, EventArgs e)
        {
            orderIdCombobox.SelectedIndex = 0;
            orderProductsCombobox.SelectedIndex = 0;
            quantityTextBox.Text = "";
            discountTextBox.Text = "";
            createOrderDetailButton.Enabled = true;
        }

        private void filterByUnitPriceTrackBar_Scroll(object sender, EventArgs e)
        {
            unitPriceValueLabel.Text = String.Format("{0:n0}", filterByUnitPriceTrackBar.Value);
        }

        private void filterByUnitsInStockTrackBar_Scroll(object sender, EventArgs e)
        {
            unitsInStockValueLabel.Text = filterByUnitsInStockTrackBar.Value.ToString();
        }

        private void searchProductButton_Click(object sender, EventArgs e)
        {
            calculatePagingProducts();
            updateProductView();
        }

        private void filterByUnitPriceTrackBar_ValueChanged(object sender, EventArgs e)
        {
            calculatePagingProducts();
            updateProductView();
        }

        private void filterByUnitsInStockTrackBar_ValueChanged(object sender, EventArgs e)
        {
            calculatePagingProducts();
            updateProductView();
        }

        private void pagingProductsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int nextPage = pagingProductsComboBox.SelectedIndex + 1;
            pagingInfoProducts.CurrentPage = nextPage;

            updateProductView();
        }

        private void previousProductsButton_Click(object sender, EventArgs e)
        {
            var currentIndex = pagingProductsComboBox.SelectedIndex;
            if (currentIndex > 0)
                pagingProductsComboBox.SelectedIndex--;
        }

        private void nextProductsButton_Click(object sender, EventArgs e)
        {
            var currentIndex = pagingProductsComboBox.SelectedIndex;
            if (currentIndex < pagingProductsComboBox.Items.Count - 1)
                pagingProductsComboBox.SelectedIndex++;
        }

        private void categoriesCombobox_SeclectionChangeCommit(object sender, EventArgs e)
        {
            calculatePagingProducts();
            updateProductView();
        }

        private void pagingCustomersCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int nextPage = pagingCustomersCombobox.SelectedIndex + 1;
            pagingInfoCustomers.CurrentPage = nextPage;

            updateCustomersView();
        }

        private void previousCustomersButton_Click(object sender, EventArgs e)
        {
            var currentIndex = pagingCustomersCombobox.SelectedIndex;
            if (currentIndex > 0)
                pagingCustomersCombobox.SelectedIndex--;
        }

        private void nextCustomersButton_Click(object sender, EventArgs e)
        {
            var currentIndex = pagingCustomersCombobox.SelectedIndex;
            if (currentIndex < pagingCustomersCombobox.Items.Count - 1)
                pagingCustomersCombobox.SelectedIndex++;
        }

        private void searchCustomerButton_Click(object sender, EventArgs e)
        {
            calculatePagingCustomers();
            updateCustomersView();
        }

        private void filterFromOrderDateDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            updateOrdersView();
        }

        private void filterToOrderDateDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            updateOrdersView();
        }

        private void filerByShipCityCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateOrdersView();
        }

        private void customerIdCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateOrdersView();
        }

        //CUSTOM DATAGRIDVIEW
        public void CustomOrdersDataGridView()
        {
            ordersDataGridView.MultiSelect = false;
            ordersDataGridView.AllowUserToAddRows = true;
            ordersDataGridView.AllowUserToDeleteRows = true;
            ordersDataGridView.AllowUserToOrderColumns = false;
            ordersDataGridView.AllowUserToResizeRows = false;
            ordersDataGridView.RowHeadersVisible = false;
            ordersDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            ordersDataGridView.AutoGenerateColumns = false;

            DataGridViewColumn columnOrderID = new DataGridViewTextBoxColumn();
            columnOrderID.ReadOnly = true;
            columnOrderID.HeaderText = "OrderID";
            columnOrderID.DataPropertyName = "OrderID";
            ordersDataGridView.Columns.Add(columnOrderID);

            DataGridViewColumn columnCustomerID = new DataGridViewTextBoxColumn();
            columnCustomerID.ReadOnly = true;
            columnCustomerID.HeaderText = "CustomerID";
            columnCustomerID.DataPropertyName = "CustomerID";
            ordersDataGridView.Columns.Add(columnCustomerID);

            DataGridViewColumn columnOrderDate = new DataGridViewTextBoxColumn();
            columnOrderDate.ReadOnly = true;
            columnOrderDate.HeaderText = "OrderDate";
            columnOrderDate.DataPropertyName = "OrderDate";
            ordersDataGridView.Columns.Add(columnOrderDate);

            DataGridViewColumn columnRequiredDate = new DataGridViewTextBoxColumn();
            columnRequiredDate.ReadOnly = true;
            columnRequiredDate.HeaderText = "RequiredDate";
            columnRequiredDate.DataPropertyName = "RequiredDate";
            ordersDataGridView.Columns.Add(columnRequiredDate);

            DataGridViewColumn columnShippedDate = new DataGridViewTextBoxColumn();
            columnShippedDate.ReadOnly = true;
            columnShippedDate.HeaderText = "ShippedDate";
            columnShippedDate.DataPropertyName = "ShippedDate";
            ordersDataGridView.Columns.Add(columnShippedDate);

            DataGridViewColumn columnFreight = new DataGridViewTextBoxColumn();
            columnFreight.ReadOnly = true;
            columnFreight.HeaderText = "Freight";
            columnFreight.DataPropertyName = "Freight";
            ordersDataGridView.Columns.Add(columnFreight);

            DataGridViewColumn columnShipName = new DataGridViewTextBoxColumn();
            columnShipName.ReadOnly = true;
            columnShipName.HeaderText = "ShipName";
            columnShipName.DataPropertyName = "ShipName";
            ordersDataGridView.Columns.Add(columnShipName);

            DataGridViewColumn columnShipAddress = new DataGridViewTextBoxColumn();
            columnShipAddress.ReadOnly = true;
            columnShipAddress.HeaderText = "ShipAddress";
            columnShipAddress.DataPropertyName = "ShipAddress";
            ordersDataGridView.Columns.Add(columnShipAddress);

            DataGridViewColumn columnShipCity = new DataGridViewTextBoxColumn();
            columnShipCity.ReadOnly = true;
            columnShipCity.HeaderText = "ShipCity";
            columnShipCity.DataPropertyName = "ShipCity";
            ordersDataGridView.Columns.Add(columnShipCity);

            DataGridViewColumn columnShipRegion = new DataGridViewTextBoxColumn();
            columnShipRegion.ReadOnly = true;
            columnShipRegion.HeaderText = "ShipRegion";
            columnShipRegion.DataPropertyName = "ShipRegion";
            ordersDataGridView.Columns.Add(columnShipRegion);

            DataGridViewColumn columnShipPostalCode = new DataGridViewTextBoxColumn();
            columnShipPostalCode.ReadOnly = true;
            columnShipPostalCode.HeaderText = "ShipPostalCode";
            columnShipPostalCode.DataPropertyName = "ShipPostalCode";
            ordersDataGridView.Columns.Add(columnShipPostalCode);

            DataGridViewColumn columnShipCountry = new DataGridViewTextBoxColumn();
            columnShipCountry.ReadOnly = true;
            columnShipCountry.HeaderText = "ShipCountry";
            columnShipCountry.DataPropertyName = "ShipCountry";
            ordersDataGridView.Columns.Add(columnShipCountry);
        }

        public void CustomOrderedProductsDataGridView()
        {
            orderedProductsDataGridView.MultiSelect = false;
            orderedProductsDataGridView.AllowUserToAddRows = true;
            orderedProductsDataGridView.AllowUserToDeleteRows = true;
            orderedProductsDataGridView.AllowUserToOrderColumns = false;
            orderedProductsDataGridView.AllowUserToResizeRows = false;
            orderedProductsDataGridView.RowHeadersVisible = false;
            orderedProductsDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            orderedProductsDataGridView.AutoGenerateColumns = false;

            DataGridViewColumn columnOrderID = new DataGridViewTextBoxColumn();
            columnOrderID.ReadOnly = true;
            columnOrderID.HeaderText = "OrderID";
            columnOrderID.DataPropertyName = "OrderID";
            orderedProductsDataGridView.Columns.Add(columnOrderID);

            DataGridViewColumn columnProductID = new DataGridViewTextBoxColumn();
            columnProductID.ReadOnly = true;
            columnProductID.HeaderText = "ProductID";
            columnProductID.DataPropertyName = "ProductID";
            orderedProductsDataGridView.Columns.Add(columnProductID);
        }

        //UTIL
        public DbDataAdapter CreateFullDataAdapter(string tableName)
        {
            var command = factory.CreateCommand();
            command.Connection = connection;
            command.CommandText = string.Format("SELECT * FROM {0}", tableName);

            var dataAdapter = factory.CreateDataAdapter();
            dataAdapter.SelectCommand = command;

            return dataAdapter;
        }

        public static byte[] ImageToByteArray(Image img, PictureBox pictureBoxCompanyLogo)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            if (pictureBoxCompanyLogo.Image != null)
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            return ms.ToArray();
        }

        public int autoGenerateIdOrder()
        {
            int orderId = northwind.Orders.Count() + 1;
            while (northwind.Orders.Select(o => o.OrderID).Contains(orderId))
            {
                orderId++;
            }
            return orderId;
        }

        public void calculatePagingProducts()
        {
            var products = northwind.Products.Where(p => p.CategoryID == int.Parse(categoriesCombobox.SelectedValue.ToString()));
            var query = from product in products
                        where product.ProductName.ToLower().Contains(searchProductTextBox.Text.ToLower())
                        && product.UnitPrice <= filterByUnitPriceTrackBar.Value
                        && product.UnitsInStock <= filterByUnitsInStockTrackBar.Value
                        select product;

            int rowPerPage = 4;
            int totalItems = query.ToList().Count;

            pagingInfoProducts = new PagingInfo()
            {
                RowsPerPage = rowPerPage,
                TotalItems = totalItems,
                TotalPages = totalItems / rowPerPage + (((totalItems % rowPerPage) == 0) ? 0 : 1),
                CurrentPage = 1
            };
            pagingProductsComboBox.DataSource = pagingInfoProducts.Pages;
        }

        public void updateProductView()
        {
            var products = northwind.Products.Where(p => p.CategoryID == int.Parse(categoriesCombobox.SelectedValue.ToString()));
            var query = from product in products
                        where product.ProductName.ToLower().Contains(searchProductTextBox.Text.ToLower())
                        && product.UnitPrice <= filterByUnitPriceTrackBar.Value
                        && product.UnitsInStock <= filterByUnitsInStockTrackBar.Value
                        select product;

            int skip = (pagingInfoProducts.CurrentPage - 1) * pagingInfoProducts.RowsPerPage;
            int take = pagingInfoProducts.RowsPerPage;

            bindingSourceProducts.DataSource = query.Skip(skip).Take(take).ToList();
        }

        public void calculatePagingCustomers()
        {
            var customers = northwind.Customers.ToList();
            var query = from customer in customers
                        where customer.ContactName.ToLower().Contains(searchCustomerTextBox.Text.ToLower())
                        select customer;

            int rowPerPage = 4;
            int totalItems = query.ToList().Count;

            pagingInfoCustomers = new PagingInfo()
            {
                RowsPerPage = rowPerPage,
                TotalItems = totalItems,
                TotalPages = totalItems / rowPerPage + (((totalItems % rowPerPage) == 0) ? 0 : 1),
                CurrentPage = 1
            };
            pagingCustomersCombobox.DataSource = pagingInfoCustomers.Pages;
        }

        public void updateCustomersView()
        {
            var customers = northwind.Customers;
            var query = from customer in customers
                        where customer.ContactName.ToLower().Contains(searchCustomerTextBox.Text.ToLower())
                        select customer;

            int skip = (pagingInfoCustomers.CurrentPage - 1) * pagingInfoCustomers.RowsPerPage;
            int take = pagingInfoCustomers.RowsPerPage;

            bindingSourceCustomers.DataSource = query.Skip(skip).Take(take).ToList();
        }

        public void updateOrdersView()
        {
            var shipCity = filerByShipCityCombobox.SelectedIndex == filerByShipCityCombobox.Items.Count - 1 ? "" : filerByShipCityCombobox.SelectedValue.ToString();

            var orders = northwind.Customers.ToList()[int.Parse(customerIdCombobox.SelectedIndex.ToString())].Orders;
            var query = from order in orders
                        where order.OrderDate >= filterFromOrderDateDateTimePicker.Value
                        && order.OrderDate <= filterToOrderDateDateTimePicker.Value
                        && order.ShipCity.ToLower().Contains(shipCity.ToLower())
                        select order;

            bindingSourceOrders.DataSource =
                sortOrdersByCombobox.SelectedIndex == 0 ? query.ToList().OrderBy(o => o.OrderDate) : query.ToList().OrderByDescending(o => o.OrderDate);
        }

        public void loadStatisticsOfProductsSoldOverMonthChart()
        {
            var O = northwind.Orders;
            var OD = northwind.Order_Details;
            var query = from o in O
                        join od in OD on o.OrderID equals od.OrderID
                        group o by o.OrderDate.Value.Month into g
                        select new
                        {
                            MONTH = g.Key,
                            COUNT_PRODUCT = g.Count()
                        };

            statisticsOfProductsSoldOverMonthChart.Series["Quantity products sold"].XValueMember = "MONTH";
            statisticsOfProductsSoldOverMonthChart.Series["Quantity products sold"].YValueMembers = "COUNT_PRODUCT";
            statisticsOfProductsSoldOverMonthChart.DataSource = query;
        }

        public void loadStatisticsOfProductsSoldOverYearChart()
        {
            var O = northwind.Orders;
            var OD = northwind.Order_Details;
            var query = from o in O
                        join od in OD on o.OrderID equals od.OrderID
                        group o by o.OrderDate.Value.Year into g
                        select new
                        {
                            YEAR = g.Key,
                            COUNT_PRODUCT = g.Count()
                        };

            statisticsOfProductsSoldOverYearChart.Series["Quantity products sold"].XValueMember = "YEAR";
            statisticsOfProductsSoldOverYearChart.Series["Quantity products sold"].YValueMembers = "COUNT_PRODUCT";
            statisticsOfProductsSoldOverYearChart.DataSource = query;
        }

        public void loadRevenueStatisticsOverMonth()
        {
            var query = from OD in northwind.Order_Details
                        group new { OD.Order, OD } by new
                        {
                            MONTH = (int?)OD.Order.OrderDate.Value.Month
                        }
                        into g
                        select new
                        {
                            MONTH = g.Key.MONTH,
                            REVENUE = (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount))
                        };

            revenueStatisticsOverMonth.Series["Revenue"].XValueMember = "MONTH";
            revenueStatisticsOverMonth.Series["Revenue"].YValueMembers = "REVENUE";
            revenueStatisticsOverMonth.DataSource = query;
        }

        public void loadRevenueStatisticsByProduct()
        {
            var query = from OD in northwind.Order_Details
                        group new { OD.Product, OD } by new
                        {
                            OD.Product.ProductName
                        }
                        into g
                        orderby
                        (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount)) descending
                        select new
                        {
                            PRODUCT = g.Key.ProductName,
                            REVENUE = (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount))
                        };

            revenueStatisticsByProduct.Series["Revenue"].XValueMember = "PRODUCT";
            revenueStatisticsByProduct.Series["Revenue"].YValueMembers = "REVENUE";
            revenueStatisticsByProduct.DataSource = query.Skip(0).Take(5);
        }

        public void loadRevenueStatisticsOverYear()
        {
            var query = from OD in northwind.Order_Details
                        group new { OD.Order, OD } by new
                        {
                            YEAR = (int?)OD.Order.OrderDate.Value.Year
                        }
                        into g
                        select new
                        {
                            YEAR = g.Key.YEAR,
                            REVENUE = (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount))
                        };

            revenueStatisticsOverYear.Series["Revenue"].XValueMember = "YEAR";
            revenueStatisticsOverYear.Series["Revenue"].YValueMembers = "REVENUE";
            revenueStatisticsOverYear.DataSource = query;
        }

        public void loadRevenueStatisticsByCustomers()
        {
            var query = from OD in northwind.Order_Details
                        group new { OD.Order.Customer, OD } by new
                        {
                            OD.Order.Customer.ContactName
                        }
                        into g
                        orderby
                        (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount)) descending
                        select new
                        {
                            CONTACTNAME = g.Key.ContactName,
                            REVENUE = (decimal?)g.Sum(p => p.OD.UnitPrice * p.OD.Quantity * (1 - (decimal)p.OD.Discount))
                        };

            revenueStatisticsByCustomersChart.Series["Revenue"].XValueMember = "CONTACTNAME";
            revenueStatisticsByCustomersChart.Series["Revenue"].YValueMembers = "REVENUE";
            revenueStatisticsByCustomersChart.DataSource = query.Skip(0).Take(5);
        }

        //BUSSINESS LOGICS
        private void loadDatabase(ConnectionStringSettings css)
        {
            factory = DbProviderFactories.GetFactory(css.ProviderName);
            connection = factory.CreateConnection();
            connection.ConnectionString = css.ConnectionString;
            loadData();
        }

        private void loadTable(string tableName, string nameDataSet)
        {
            var command = factory.CreateCommand();
            command.Connection = connection;
            command.CommandText = string.Format("SELECT * FROM [{0}]", tableName);

            var dataAdapter = factory.CreateDataAdapter();
            dataAdapter.SelectCommand = command;

            dataSet.EnforceConstraints = false;

            dataAdapter.Fill(dataSet, nameDataSet);
        }

        private void loadData()
        {
            loadTable(tableCategoriesName, tableCategoriesName);
            loadTable(tableProductsName, tableProductsName);

            dataSet.Relations.Add(tableCategories_Products_RelationnName, dataSet.Tables[tableCategoriesName].Columns["CategoryID"],
               dataSet.Tables[tableProductsName].Columns["CategoryID"]);

            //Categories
            bindingSourceCategories.DataSource = dataSet;
            bindingSourceCategories.DataMember = tableCategoriesName;
            categoriesCombobox.DataSource = bindingSourceCategories;
            categoriesCombobox.DisplayMember = "CategoryName";
            categoriesCombobox.ValueMember = "CategoryID";
            categoryNameTextBox.DataBindings.Add("Text", bindingSourceCategories, "CategoryName", false, DataSourceUpdateMode.Never);
            picturePictureBox.DataBindings.Add("Image", bindingSourceCategories, "Picture", true, DataSourceUpdateMode.Never); //using dataSet because not solve
            categoryDesscriptionTextBox.DataBindings.Add("Text", bindingSourceCategories, "Description", false, DataSourceUpdateMode.Never);

            //Categories combobox (Products)
            categoryCombobox.DataSource = bindingSourceCategories;
            categoryCombobox.DisplayMember = "CategoryName";
            categoryCombobox.ValueMember = "CategoryID";

            //Products
            bindingSourceProducts.DataSource = bindingSourceCategories;
            bindingSourceProducts.DataMember = tableCategories_Products_RelationnName;
            productsDataGridView.DataSource = bindingSourceProducts;
            productsDataGridView.AutoGenerateColumns = false;
            productIdTextBox.DataBindings.Add("Text", bindingSourceProducts, "ProductID", false, DataSourceUpdateMode.Never);
            productNameTextBox.DataBindings.Add("Text", bindingSourceProducts, "ProductName", false, DataSourceUpdateMode.Never);
            categoryCombobox.DataBindings.Add("SelectedValue", bindingSourceProducts, "CategoryID", false, DataSourceUpdateMode.Never);
            quantityPerUnitTextBox.DataBindings.Add("Text", bindingSourceProducts, "QuantityPerUnit", false, DataSourceUpdateMode.Never);
            unitPriceTextBox.DataBindings.Add("Text", bindingSourceProducts, "UnitPrice", false, DataSourceUpdateMode.Never);
            UnitsInStockTextBox.DataBindings.Add("Text", bindingSourceProducts, "UnitsInStock", false, DataSourceUpdateMode.Never);
            UnitsOnOrderTextBox.DataBindings.Add("Text", bindingSourceProducts, "UnitsOnOrder", false, DataSourceUpdateMode.Never);
            ReorderLevelTextBox.DataBindings.Add("Text", bindingSourceProducts, "ReorderLevel", false, DataSourceUpdateMode.Never);
            discontinuedCheckBox.DataBindings.Add("Checked", bindingSourceProducts, "Discontinued", false, DataSourceUpdateMode.Never);
            calculatePagingProducts();

            //Customers
            customersDataGridView.DataSource = bindingSourceCustomers;
            customerIdTextBox.DataBindings.Add("Text", bindingSourceCustomers, "CustomerID", false, DataSourceUpdateMode.Never);
            companyNameTextBox.DataBindings.Add("Text", bindingSourceCustomers, "CompanyName", false, DataSourceUpdateMode.Never);
            contactNameTextBox.DataBindings.Add("Text", bindingSourceCustomers, "ContactName", false, DataSourceUpdateMode.Never);
            contactTitleTextBox.DataBindings.Add("Text", bindingSourceCustomers, "ContactTitle", false, DataSourceUpdateMode.Never);
            addressTextBox.DataBindings.Add("Text", bindingSourceCustomers, "Address", false, DataSourceUpdateMode.Never);
            cityTextBox.DataBindings.Add("Text", bindingSourceCustomers, "City", false, DataSourceUpdateMode.Never);
            regionTextBox.DataBindings.Add("Text", bindingSourceCustomers, "Region", false, DataSourceUpdateMode.Never);
            postalCodeTextBox.DataBindings.Add("Text", bindingSourceCustomers, "PostalCode", false, DataSourceUpdateMode.Never);
            countryTextBox.DataBindings.Add("Text", bindingSourceCustomers, "Country", false, DataSourceUpdateMode.Never);
            phoneTextBox.DataBindings.Add("Text", bindingSourceCustomers, "Phone", false, DataSourceUpdateMode.Never);
            faxTextBox.DataBindings.Add("Text", bindingSourceCustomers, "Fax", false, DataSourceUpdateMode.Never);
            calculatePagingCustomers();

            //Orders
            bindingSourceOrders.DataSource = bindingSourceCustomers;
            bindingSourceOrders.DataMember = "Orders";
            customerIdCombobox.DataSource = northwind.Customers;
            customerIdCombobox.DisplayMember = "CustomerID";
            customerIdCombobox.ValueMember = "CustomerID";
            ordersDataGridView.DataSource = bindingSourceOrders;
            CustomOrdersDataGridView();
            sortOrdersByCombobox.SelectedIndex = 0;
            filerByShipCityCombobox.DataSource = northwind.Orders.Select(o => o.ShipCity).Distinct().Append("All").ToList();
            filerByShipCityCombobox.SelectedIndex = filerByShipCityCombobox.Items.Count - 1;
            filterFromOrderDateDateTimePicker.Value = new DateTime(2020, 01, 01);
            filterToOrderDateDateTimePicker.Value = new DateTime(2021, 12, 31);

            //Ordered Product
            bindingSourceOrderDetails.DataSource = bindingSourceOrders;
            bindingSourceOrderDetails.DataMember = "Order_Details";
            bindingSourceOrderedProducts.DataSource = bindingSourceOrderDetails;
            bindingSourceOrderedProducts.DataMember = "Product";
            CustomOrderedProductsDataGridView();
            orderedProductsDataGridView.DataSource = bindingSourceOrderedProducts.DataSource;

            //Create Order
            orderCustomerDataGridView.DataSource = northwind.Customers;
            orderIdTextBox.Text = autoGenerateIdOrder().ToString();
            orderCustomerIdTextBox.DataBindings.Add("Text", orderCustomerDataGridView.DataSource, "CustomerID", false, DataSourceUpdateMode.Never);

            //Create Order Detail
            orderIdCombobox.DataSource = bindingSourceOrders;
            orderIdCombobox.DisplayMember = "OrderID";
            orderIdCombobox.ValueMember = "OrderID";
            orderProductsCombobox.DataSource = northwind.Products;
            orderProductsCombobox.DisplayMember = "ProductName";
            orderProductsCombobox.ValueMember = "ProductID";
            orderUnitPriceTextBox.DataBindings.Add("Text", orderProductsCombobox.DataSource, "UnitPrice", false, DataSourceUpdateMode.Never);

            //Statitics of Products sold over month
            loadStatisticsOfProductsSoldOverMonthChart();

            //Statitics of Products sold over year
            loadStatisticsOfProductsSoldOverYearChart();

            //Revenue statistics by Product
            loadRevenueStatisticsByProduct();

            //Revenue statistics over month
            loadRevenueStatisticsOverMonth();

            //Revenue statistics over year
            loadRevenueStatisticsOverYear();

            //Revenue statistics by Customers
            loadRevenueStatisticsByCustomers();
        }

        private void addCategory(string categoryName, string description, Byte[] picture)
        {
            int categoryId = northwind.Categories.Count() + 1;
            while (northwind.Categories.Select(c => c.CategoryID).Contains(categoryId))
            {
                categoryId++;
            }

            var category = new Category();
            category.CategoryID = categoryId;
            category.CategoryName = categoryName;
            category.Description = description;
            category.Picture = picture;

            northwind.Categories.InsertOnSubmit(category);
            northwind.SubmitChanges();
        }

        private void updateCategory(int categoryId, string categoryName, string description, byte[] picture)
        {
            var category = northwind.Categories.Single(c => c.CategoryID == categoryId);
            category.CategoryName = categoryName;
            category.Description = description;
            category.Picture = picture;

            northwind.SubmitChanges();
        }

        private void deleteCategory(int categoryId)
        {
            var category = northwind.Categories.Single(c => c.CategoryID == categoryId);
            northwind.Categories.DeleteOnSubmit(category);
            northwind.SubmitChanges();
        }

        private void addProducts(string productName, int categoryId, string quantityPerUnit, decimal unitPrice, short unitsInStock, short unitsOnOrder, short reorderLevel, bool discontinued)
        {
            int productId = northwind.Products.Count() + 1;
            while (northwind.Products.Select(p => p.ProductID).Contains(productId))
            {
                productId++;
            }

            var product = new Product();
            product.ProductID = productId;
            product.ProductName = productName;
            product.CategoryID = categoryId;
            product.QuantityPerUnit = quantityPerUnit;
            product.UnitPrice = unitPrice;
            product.UnitsInStock = unitsInStock;
            product.UnitsOnOrder = unitsOnOrder;
            product.ReorderLevel = reorderLevel;
            product.Discontinued = discontinued;

            northwind.Products.InsertOnSubmit(product);
            northwind.SubmitChanges();
        }

        private void updateProducts(int productId, string productName, int categoryId, string quantityPerUnit, decimal unitPrice, short unitsInStock, short unitsOnOrder, short reorderLevel, bool discontinued)
        {
            var product = northwind.Products.Single(p => p.ProductID == productId);
            product.ProductName = productName;
            product.CategoryID = categoryId;
            product.QuantityPerUnit = quantityPerUnit;
            product.UnitPrice = unitPrice;
            product.UnitsInStock = unitsInStock;
            product.UnitsOnOrder = unitsOnOrder;
            product.ReorderLevel = reorderLevel;
            product.Discontinued = discontinued;

            northwind.SubmitChanges();
        }

        private void deleteProduct(int productId)
        {
            var product = northwind.Products.Single(p => p.ProductID == productId);
            northwind.Products.DeleteOnSubmit(product);
            northwind.SubmitChanges();
        }

        private void addCustomer(string companyName, string contactName, string contactTitle, string address, string city, string region, string postalCode, string country, string phone, string fax)
        {
            string customerId = "";
            for (int i = 0; i < companyName.Length; i++)
            {
                if (companyName[i] == ' ') continue;
                if (i == 5) break;
                customerId += companyName[i];
                try
                {
                    if (northwind.Customers.Select(c => c.CustomerID).Contains(customerId))
                        customerId = customerId.Replace(companyName[i], (char)(int)(companyName[i] + 1));
                }
                catch (Exception) { }
            }

            var customer = new Customer();
            customer.CustomerID = customerId;
            customer.CompanyName = companyName;
            customer.ContactName = contactName;
            customer.ContactTitle = contactTitle;
            customer.Address = address;
            customer.City = city;
            customer.Region = region;
            customer.PostalCode = postalCode;
            customer.Country = country;
            customer.Phone = phone;
            customer.Fax = fax;

            northwind.Customers.InsertOnSubmit(customer);
            northwind.SubmitChanges();
        }

        private void updateCustomer(string customerId, string companyName, string contactName, string contactTitle, string address, string city, string region, string postalCode, string country, string phone, string fax)
        {
            var customer = northwind.Customers.Single(c => c.CustomerID == customerId);
            customer.CompanyName = companyName;
            customer.ContactName = contactName;
            customer.ContactTitle = contactTitle;
            customer.Address = address;
            customer.City = city;
            customer.Region = region;
            customer.PostalCode = postalCode;
            customer.Country = country;
            customer.Phone = phone;
            customer.Fax = fax;

            northwind.SubmitChanges();
        }

        private void deleteCustomer(string customerId)
        {
            var customer = northwind.Customers.Single(c => c.CustomerID == customerId);
            northwind.Customers.DeleteOnSubmit(customer);
            northwind.SubmitChanges();
        }

        private void addOrders(int orderId, string customerId, DateTime orderDate, DateTime requiredDate, DateTime shippedDate, decimal freight, string shipName, string shipAddress, string shipCity, string shipRegion, string shipPostalCode, string shipCountry)
        {
            var order = new Order();
            order.OrderID = orderId;
            order.CustomerID = customerId;
            order.OrderDate = orderDate;
            order.RequiredDate = requiredDate;
            order.ShippedDate = shippedDate;
            order.Freight = freight;
            order.ShipName = shipName;
            order.ShipAddress = shipAddress;
            order.ShipCity = shipCity;
            order.ShipRegion = shipRegion;
            order.ShipPostalCode = shipPostalCode;
            order.ShipCountry = shipCountry;

            northwind.Orders.InsertOnSubmit(order);
            northwind.SubmitChanges();
        }

        private void addOrderDetail(int orderId, int productId, decimal unitPrice, short quantity, float discount)
        {
            var orderDetails = new Order_Detail();
            orderDetails.OrderID = orderId;
            orderDetails.ProductID = productId;
            orderDetails.UnitPrice = unitPrice;
            orderDetails.Quantity = quantity;
            orderDetails.Discount = discount;

            northwind.Order_Details.InsertOnSubmit(orderDetails);
            northwind.SubmitChanges();
        }
    }
}