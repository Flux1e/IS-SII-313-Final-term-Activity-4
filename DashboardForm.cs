using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;


namespace Inventory
{
    public partial class DashboardForm : Form
    {

        private int selectedProductID = 0;

        public DashboardForm(SignInForm login)
        {
            InitializeComponent();

            btnLogOut.Click += btnLogOut_Click;
            btnAdd.Click += btnAdd_Click;
            btnUpdate.Click += btnUpdate_Click;
            btnDelete.Click += btnDelete_Click;
            btnClear.Click += btnClear_Click;
            btnRefresh.Click += btnRefresh_Click;

            cmbCategory.DropDownStyle = ComboBoxStyle.DropDownList;

            txtSearch.TextChanged += txtSearch_TextChanged;

            dgvProducts.CellClick += dgvProducts_CellClick;

            LoadProducts();
            LoadDashboardStats();

        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if(selectedProductID == 0)
            {
                MessageBox.Show("Please select a product to delete.");
                return;
            }
            DialogResult result = MessageBox.Show(
                "Are you sure you want to delete this product?",
                "Confirm Deletion",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if(result == DialogResult.Yes)
            {
                try
                {
                    using (OleDbConnection con = Database.GetConnection())
                    {
                        con.Open();
                        string query = "DELETE FROM Products WHERE ProductID = ?";
                        using (OleDbCommand cmd = new OleDbCommand(query, con))
                        {
                            cmd.Parameters.Add("?", OleDbType.Integer).Value = selectedProductID;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    MessageBox.Show("Product deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    clearField();
                    LoadProducts();
                    LoadDashboardStats();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Deleting Product:" + ex.Message);
                }
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if(selectedProductID == 0)
            {
                MessageBox.Show("Please select a product to update.");
                return;
            }
            if(!ValidateProducts())
                return;

            try
            {
                using (OleDbConnection con = Database.GetConnection())
                {
                    con.Open();
                    string query = @"UPDATE Products
    SET ProductName = ?, Category = ?, Description = ?,
    UnitPrice = ?, Quantity = ?, [Reorder Level] = ?
    WHERE ID = ?";
                    using (OleDbCommand cmd = new OleDbCommand(query, con))
                    {
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value = txtProductName.Text.Trim();
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value = cmbCategory.Text;
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value = txtDescription.Text.Trim();
                        cmd.Parameters.Add("?", OleDbType.Currency).Value = decimal.Parse(txtUnitPrice.Text);
                        cmd.Parameters.Add("?", OleDbType.Integer).Value = int.Parse(txtQuantity.Text);
                        cmd.Parameters.Add("?", OleDbType.Integer).Value = int.Parse(txtReorderLevel.Text);
                        cmd.Parameters.Add("?", OleDbType.Integer).Value = selectedProductID;
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Product Updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                clearField();
                LoadProducts();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Updating Product:" + ex.Message);
            }
        }



        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateProducts())
                return;

            try
            {
                using (OleDbConnection con = Database.GetConnection())
                {
                    con.Open();

                    string query = @"INSERT INTO Products
                (ProductName, Category, Description, UnitPrice, Quantity, [Reorder Level])
                VALUES (?, ?, ?, ?, ?, ?)";

                    using (OleDbCommand cmd = new OleDbCommand(query, con))
                    {
                        cmd.Parameters.Add("?", OleDbType.VarChar).Value =
                            txtProductName.Text.Trim();

                        cmd.Parameters.Add("?", OleDbType.VarChar).Value =
                            cmbCategory.Text;

                        cmd.Parameters.Add("?", OleDbType.VarChar).Value =
                            txtDescription.Text.Trim();

                        cmd.Parameters.Add("?", OleDbType.Currency).Value =
                            decimal.Parse(txtUnitPrice.Text);

                        cmd.Parameters.Add("?", OleDbType.Integer).Value =
                            int.Parse(txtQuantity.Text);

                        cmd.Parameters.Add("?", OleDbType.Integer).Value =
                            int.Parse(txtReorderLevel.Text);

                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Product added successfully!",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                clearField();
                LoadProducts();
                LoadDashboardStats();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Adding Product: " + ex.Message);
            }
        }

        private void clearField()
        {
            txtProductName.Clear();
            cmbCategory.SelectedIndex = -1;
            txtDescription.Clear();
            txtUnitPrice.Clear();
            txtQuantity.Clear();
            txtReorderLevel.Clear();

            selectedProductID = 0;
            dgvProducts.ClearSelection();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
        }

        private void dgvProducts_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvProducts.Rows[e.RowIndex];

                selectedProductID = Convert.ToInt32(row.Cells["ProductID"].Value);

                txtProductName.Text = row.Cells["ProductName"].Value.ToString();
                cmbCategory.Text = row.Cells["Category"].Value.ToString();
                txtDescription.Text = row.Cells["Description"].Value.ToString();
                txtUnitPrice.Text = row.Cells["UnitPrice"].Value.ToString();
                txtQuantity.Text = row.Cells["Quantity"].Value.ToString();
                txtReorderLevel.Text = row.Cells["ReorderLevel"].Value.ToString();
            }
        }
         

        private void LoadDashboardStats()
        {
            try
            {
                using (OleDbConnection con = Database.GetConnection())
                {
                    con.Open();

                    string query = @"
SELECT
    COUNT(*) AS TotalProducts,

    SUM(IIf(
        IIf(IsNull(Quantity), 0, Quantity) >
        IIf(IsNull([Reorder Level]), 0, [Reorder Level]),
        1, 0
    )) AS InStock,

    SUM(IIf(
        IIf(IsNull(Quantity), 0, Quantity) > 0
        AND
        IIf(IsNull(Quantity), 0, Quantity) <=
        IIf(IsNull([Reorder Level]), 0, [Reorder Level]),
        1, 0
    )) AS LowStock,

    SUM(IIf(
        IIf(IsNull(Quantity), 0, Quantity) = 0,
        1, 0
    )) AS OutOfStock

FROM Products;
";
                    using (OleDbCommand cmd = new OleDbCommand(query, con))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            lblTotal.Text = reader["TotalProducts"].ToString();
                            lblInStock.Text = reader["InStock"].ToString();
                            lblLowStock.Text = reader["LowStock"].ToString();
                            lblOutOfStock.Text = reader["OutOfStock"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Statistics Error:" + ex.Message);
            }
        }



        private void LoadProducts()
        {
            try
            {
                using (OleDbConnection con = Database.GetConnection())
                {
                    con.Open();

                    string query = "SELECT * FROM Products ORDER BY ID DESC";

                    using (OleDbDataAdapter adapter =
                            new OleDbDataAdapter(query, con))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        dgvProducts.DataSource = dt;
                        dgvProducts.AutoSizeColumnsMode =
                            DataGridViewAutoSizeColumnsMode.Fill;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to load products: " + ex.Message);
            }
        }



        private bool ValidateProducts()
        {

            if (string.IsNullOrWhiteSpace(txtProductName.Text))
            {
                MessageBox.Show("Please enter the Product Name.");
                txtProductName.Focus();
                return false;
            }
            if (cmbCategory.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Category.");
                cmbCategory.Focus();
                return false;
            }

            decimal price;
            if (!decimal.TryParse(txtUnitPrice.Text, out price) || price <= 0)
            {
                MessageBox.Show("Please enter a valid Unit Price.");
                txtUnitPrice.Focus();
                return false;
            }

            int quantity;
            if (!int.TryParse(txtQuantity.Text, out quantity) || quantity < 0)
            {
                MessageBox.Show("Please enter a valid Quantity.");
                txtQuantity.Focus();
                return false;
            }

            int reorder;
            if (!int.TryParse(txtReorderLevel.Text, out reorder) || reorder < 0)
            {
                MessageBox.Show("Please enter a valid Reorder Level.");
                txtReorderLevel.Focus();
                return false;
            }
            return true;
        }

        private void btnadd_Click(object sender, EventArgs e)
        {

        }

        private void btnLogOut_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Are you sure you want to log out?",
                "Confirm Logout",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                Application.Restart();
            }
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            clearField();
        }

        private void InitializeComponent()
        {
            label12 = new Label();
            label10 = new Label();
            label17 = new Label();
            label4 = new Label();
            label5 = new Label();
            panel4 = new Panel();
            lblOutOfStock = new Label();
            label7 = new Label();
            label3 = new Label();
            label6 = new Label();
            panel1 = new Panel();
            lblTotal = new Label();
            label2 = new Label();
            label1 = new Label();
            btnLogOut = new Button();
            dgvProducts = new DataGridView();
            lblInStock = new Label();
            txtSearch = new TextBox();
            lblLowStock = new Label();
            btnUpdate = new Button();
            btnClear = new Button();
            btnDelete = new Button();
            btnAdd = new Button();
            txtReorderLevel = new TextBox();
            txtQuantity = new TextBox();
            label11 = new Label();
            cmbCategory = new ComboBox();
            txtDescription = new TextBox();
            txtUnitPrice = new TextBox();
            txtProductName = new TextBox();
            btnRefresh = new Button();
            label18 = new Label();
            panel3 = new Panel();
            label8 = new Label();
            panel2 = new Panel();
            label9 = new Label();
            panel4.SuspendLayout();
            panel1.SuspendLayout();
            ((ISupportInitialize)dgvProducts).BeginInit();
            panel3.SuspendLayout();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // label12
            // 
            label12.AutoSize = true;
            label12.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label12.Location = new Point(60, 347);
            label12.Name = "label12";
            label12.Size = new Size(87, 20);
            label12.TabIndex = 96;
            label12.Text = "Description";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label10.Location = new Point(527, 347);
            label10.Name = "label10";
            label10.Size = new Size(103, 20);
            label10.TabIndex = 95;
            label10.Text = "Reorder Level";
            // 
            // label17
            // 
            label17.AutoSize = true;
            label17.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label17.Location = new Point(22, 461);
            label17.Name = "label17";
            label17.Size = new Size(100, 21);
            label17.TabIndex = 97;
            label17.Text = "Product List";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label4.Location = new Point(521, 263);
            label4.Name = "label4";
            label4.Size = new Size(76, 20);
            label4.TabIndex = 85;
            label4.Text = "Unit Price";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label5.Location = new Point(294, 263);
            label5.Name = "label5";
            label5.Size = new Size(72, 20);
            label5.TabIndex = 87;
            label5.Text = "Category";
            // 
            // panel4
            // 
            panel4.BackColor = Color.WhiteSmoke;
            panel4.Controls.Add(lblOutOfStock);
            panel4.Controls.Add(label7);
            panel4.Location = new Point(694, 97);
            panel4.Name = "panel4";
            panel4.Size = new Size(159, 135);
            panel4.TabIndex = 91;
            // 
            // lblOutOfStock
            // 
            lblOutOfStock.AutoSize = true;
            lblOutOfStock.Font = new Font("Segoe UI", 24F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblOutOfStock.Location = new Point(61, 62);
            lblOutOfStock.Name = "lblOutOfStock";
            lblOutOfStock.Size = new Size(38, 45);
            lblOutOfStock.TabIndex = 10;
            lblOutOfStock.Text = "0";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label7.Location = new Point(28, 10);
            label7.Name = "label7";
            label7.Size = new Size(104, 21);
            label7.TabIndex = 4;
            label7.Text = "Out of Stock";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label3.Location = new Point(20, 10);
            label3.Name = "label3";
            label3.Size = new Size(119, 21);
            label3.TabIndex = 0;
            label3.Text = "Total Products";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label6.Location = new Point(60, 263);
            label6.Name = "label6";
            label6.Size = new Size(108, 20);
            label6.TabIndex = 89;
            label6.Text = "Product Name";
            // 
            // panel1
            // 
            panel1.BackColor = Color.WhiteSmoke;
            panel1.Controls.Add(lblTotal);
            panel1.Controls.Add(label3);
            panel1.Location = new Point(82, 97);
            panel1.Name = "panel1";
            panel1.Size = new Size(159, 135);
            panel1.TabIndex = 90;
            // 
            // lblTotal
            // 
            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Segoe UI", 21.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblTotal.Location = new Point(63, 62);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(34, 40);
            lblTotal.TabIndex = 11;
            lblTotal.Text = "0";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 15.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label2.Location = new Point(5, 46);
            label2.Name = "label2";
            label2.Size = new Size(227, 30);
            label2.TabIndex = 88;
            label2.Text = "Product Management";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 14.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.Location = new Point(5, 6);
            label1.Name = "label1";
            label1.Size = new Size(99, 25);
            label1.TabIndex = 84;
            label1.Text = "Welcome!";
            // 
            // btnLogOut
            // 
            btnLogOut.BackColor = Color.SaddleBrown;
            btnLogOut.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnLogOut.ForeColor = SystemColors.ButtonFace;
            btnLogOut.Location = new Point(846, 3);
            btnLogOut.Name = "btnLogOut";
            btnLogOut.Size = new Size(101, 32);
            btnLogOut.TabIndex = 86;
            btnLogOut.Text = "Log out";
            btnLogOut.UseVisualStyleBackColor = false;
            // 
            // dgvProducts
            // 
            dgvProducts.AllowUserToOrderColumns = true;
            dgvProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvProducts.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvProducts.Location = new Point(22, 499);
            dgvProducts.Name = "dgvProducts";
            dgvProducts.Size = new Size(896, 254);
            dgvProducts.TabIndex = 94;
            // 
            // lblInStock
            // 
            lblInStock.AutoSize = true;
            lblInStock.Font = new Font("Segoe UI", 21.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblInStock.Location = new Point(62, 67);
            lblInStock.Name = "lblInStock";
            lblInStock.Size = new Size(34, 40);
            lblInStock.TabIndex = 10;
            lblInStock.Text = "0";
            // 
            // txtSearch
            // 
            txtSearch.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtSearch.Location = new Point(716, 462);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(196, 27);
            txtSearch.TabIndex = 111;
            // 
            // lblLowStock
            // 
            lblLowStock.AutoSize = true;
            lblLowStock.Font = new Font("Segoe UI", 21.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblLowStock.Location = new Point(59, 67);
            lblLowStock.Name = "lblLowStock";
            lblLowStock.Size = new Size(34, 40);
            lblLowStock.TabIndex = 9;
            lblLowStock.Text = "0";
            // 
            // btnUpdate
            // 
            btnUpdate.BackColor = Color.DarkTurquoise;
            btnUpdate.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnUpdate.Location = new Point(775, 301);
            btnUpdate.Name = "btnUpdate";
            btnUpdate.Size = new Size(114, 33);
            btnUpdate.TabIndex = 109;
            btnUpdate.Text = "Update";
            btnUpdate.UseVisualStyleBackColor = false;
            // 
            // btnClear
            // 
            btnClear.BackColor = Color.Plum;
            btnClear.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnClear.ForeColor = Color.Black;
            btnClear.Location = new Point(775, 375);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(114, 33);
            btnClear.TabIndex = 108;
            btnClear.Text = "Clear";
            btnClear.UseVisualStyleBackColor = false;
            // 
            // btnDelete
            // 
            btnDelete.BackColor = Color.Tomato;
            btnDelete.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnDelete.Location = new Point(775, 340);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(114, 27);
            btnDelete.TabIndex = 107;
            btnDelete.Text = "Delete";
            btnDelete.UseVisualStyleBackColor = false;
            // 
            // btnAdd
            // 
            btnAdd.BackColor = Color.SpringGreen;
            btnAdd.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnAdd.ForeColor = Color.Black;
            btnAdd.Location = new Point(775, 263);
            btnAdd.Name = "btnAdd";
            btnAdd.Size = new Size(114, 32);
            btnAdd.TabIndex = 106;
            btnAdd.Text = "Add Product";
            btnAdd.UseVisualStyleBackColor = false;
            // 
            // txtReorderLevel
            // 
            txtReorderLevel.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtReorderLevel.Location = new Point(521, 370);
            txtReorderLevel.Name = "txtReorderLevel";
            txtReorderLevel.Size = new Size(176, 27);
            txtReorderLevel.TabIndex = 105;
            // 
            // txtQuantity
            // 
            txtQuantity.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtQuantity.Location = new Point(294, 371);
            txtQuantity.Name = "txtQuantity";
            txtQuantity.Size = new Size(176, 27);
            txtQuantity.TabIndex = 104;
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label11.Location = new Point(294, 347);
            label11.Name = "label11";
            label11.Size = new Size(68, 20);
            label11.TabIndex = 103;
            label11.Text = "Quantity";
            // 
            // cmbCategory
            // 
            cmbCategory.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            cmbCategory.FormattingEnabled = true;
            cmbCategory.Items.AddRange(new object[] { "GPU", "CPU", "RAM", "STORAGE", "PSU", "FANS", "Peripherals" });
            cmbCategory.Location = new Point(294, 287);
            cmbCategory.Name = "cmbCategory";
            cmbCategory.Size = new Size(178, 28);
            cmbCategory.TabIndex = 102;
            // 
            // txtDescription
            // 
            txtDescription.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtDescription.Location = new Point(60, 371);
            txtDescription.Name = "txtDescription";
            txtDescription.Size = new Size(186, 27);
            txtDescription.TabIndex = 101;
            // 
            // txtUnitPrice
            // 
            txtUnitPrice.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtUnitPrice.Location = new Point(521, 288);
            txtUnitPrice.Name = "txtUnitPrice";
            txtUnitPrice.Size = new Size(176, 27);
            txtUnitPrice.TabIndex = 100;
            // 
            // txtProductName
            // 
            txtProductName.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtProductName.Location = new Point(60, 287);
            txtProductName.Name = "txtProductName";
            txtProductName.Size = new Size(181, 27);
            txtProductName.TabIndex = 99;
            // 
            // btnRefresh
            // 
            btnRefresh.BackColor = Color.SaddleBrown;
            btnRefresh.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnRefresh.ForeColor = Color.WhiteSmoke;
            btnRefresh.Location = new Point(145, 457);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(86, 28);
            btnRefresh.TabIndex = 110;
            btnRefresh.Text = "Refresh";
            btnRefresh.UseVisualStyleBackColor = false;
            // 
            // label18
            // 
            label18.AutoSize = true;
            label18.Font = new Font("Segoe UI Semibold", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label18.Location = new Point(591, 465);
            label18.Name = "label18";
            label18.Size = new Size(119, 20);
            label18.TabIndex = 98;
            label18.Text = "Search Products";
            // 
            // panel3
            // 
            panel3.BackColor = Color.WhiteSmoke;
            panel3.Controls.Add(lblLowStock);
            panel3.Controls.Add(label8);
            panel3.Location = new Point(487, 97);
            panel3.Name = "panel3";
            panel3.Size = new Size(159, 135);
            panel3.TabIndex = 92;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label8.Location = new Point(34, 10);
            label8.Name = "label8";
            label8.Size = new Size(87, 21);
            label8.TabIndex = 5;
            label8.Text = "Low Stock";
            // 
            // panel2
            // 
            panel2.BackColor = Color.WhiteSmoke;
            panel2.Controls.Add(lblInStock);
            panel2.Controls.Add(label9);
            panel2.Location = new Point(283, 97);
            panel2.Name = "panel2";
            panel2.Size = new Size(159, 135);
            panel2.TabIndex = 93;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label9.Location = new Point(43, 10);
            label9.Name = "label9";
            label9.Size = new Size(71, 21);
            label9.TabIndex = 6;
            label9.Text = "In Stock";
            // 
            // DashboardForm
            // 
            BackColor = Color.Bisque;
            ClientSize = new Size(953, 756);
            Controls.Add(label12);
            Controls.Add(label10);
            Controls.Add(label17);
            Controls.Add(label4);
            Controls.Add(label5);
            Controls.Add(panel4);
            Controls.Add(label6);
            Controls.Add(panel1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(btnLogOut);
            Controls.Add(dgvProducts);
            Controls.Add(txtSearch);
            Controls.Add(btnUpdate);
            Controls.Add(btnClear);
            Controls.Add(btnDelete);
            Controls.Add(btnAdd);
            Controls.Add(txtReorderLevel);
            Controls.Add(txtQuantity);
            Controls.Add(label11);
            Controls.Add(cmbCategory);
            Controls.Add(txtDescription);
            Controls.Add(txtUnitPrice);
            Controls.Add(txtProductName);
            Controls.Add(btnRefresh);
            Controls.Add(label18);
            Controls.Add(panel3);
            Controls.Add(panel2);
            Name = "DashboardForm";
            Load += DashboardForm_Load;
            panel4.ResumeLayout(false);
            panel4.PerformLayout();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((ISupportInitialize)dgvProducts).EndInit();
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            ResumeLayout(false);
            PerformLayout();

        }

        private void DashboardForm_Load(object sender, EventArgs e)
        {

        }
    }
}
