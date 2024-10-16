using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using static Libreria.Base_Estudiante;

namespace Libreria    
{   
    public partial class Base_Estudiante : Form
    {      
        public Base_Estudiante()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }        

        private void Base_Estudiante_Load(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";
            LoadExcelData(excelFilePath);
            panel4.Hide();
        }

        private void LoadExcelData(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                
                var worksheet = package.Workbook.Worksheets[0]; // Primer hoja
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;


                // Agregar columnas al DataTable desde la columna C (3) hasta la columna H (8)
                for (int col = 1; col <= 6; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text); // Usar la fila 4 para los encabezados
                }

                // Rellena la tabla comenzando desde la fila 5
                // Cambia rowCount por el rango específico si es necesario
                for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 5
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= 6; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text; // Ajusta el índice para el DataTable
                    }
                    dataTable.Rows.Add(newRow);
                }
            }
            // Asigna el DataTable al DataGridView
            dataGridView1.DataSource = dataTable;
        }

        private void LoadTesis(string filePath)
        {
            var dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {

                var worksheet = package.Workbook.Worksheets[1]; // Primer hoja
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;


                // Agregar columnas al DataTable desde la columna C (3) hasta la columna H (8)
                for (int col = 1; col <= 5; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text); // Usar la fila 4 para los encabezados
                }

                // Rellena la tabla comenzando desde la fila 5
                // Cambia rowCount por el rango específico si es necesario
                for (int row = 2; row <= rowCount; row++) // Comienza desde la fila 5
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= 5; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text; // Ajusta el índice para el DataTable
                    }
                    dataTable.Rows.Add(newRow);
                }
            }
            // Asigna el DataTable al DataGridView
            dataGridView1.DataSource = dataTable;
        }

        private void retirar_Click(object sender, EventArgs e)
        {
            retirar sexo = new retirar();
            sexo.MostarPanel1();            
            sexo.Show();
        }

        private void ingresar_Click(object sender, EventArgs e)
        {
            agregar sexo = new agregar();
            sexo.MostrarPanelLibro();
            sexo.Show();
        }
               
        private void lIBROSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";
            LoadExcelData(excelFilePath);
            dataGridView1.Show();
            pnlibro.Show();
            pntesis.Hide();
            panel3.Show();
            panel2.Hide();
            panel4.Hide();
            button1.Show();
            button5.Hide();
            Titulo.Checked = true;
        }

        private void tESISDEGRADOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";
            LoadTesis(excelFilePath);
            dataGridView2.Show();
            pntesis.Show();
            pnlibro.Hide();
            panel3.Hide();
            panel2.Show();
            panel4.Show();
            button1.Hide();
            button5.Show();
            radioButton9.Checked = true;
        }
        
        private void regresar_Click(object sender, EventArgs e)
        {
            retirar sexo = new retirar();                       
            sexo.MostrarPanel3();
            sexo.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";
            LoadExcelData(excelFilePath);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo de búsqueda no esté vacío
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {                
                LoadExcelData(excelFilePath);
                return;
            }

            string searchTerm = textBox1.Text; // Lo que se buscará
            int searchColumn = 0; // Inicializa con la columna de búsqueda

            // Determinar cuál RadioButton está seleccionado para definir la columna de búsqueda
            if (Titulo.Checked) searchColumn = 1; // Primera columna
            else if (radioButton2.Checked) searchColumn = 2; // Segunda columna
            else if (radioButton3.Checked) searchColumn = 4; // Tercera columna
            else if (radioButton4.Checked) searchColumn = 5; // Cuarta columna
            else if (radioButton5.Checked) searchColumn = 6; // Quinta columna

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja
                DataTable dt = new DataTable();

                // Crear columnas en el DataTable (Ajusta las columnas según tus necesidades)
                dt.Columns.Add("Título", typeof(string));
                dt.Columns.Add("Autor", typeof(string));                
                dt.Columns.Add("ISBN", typeof(string));
                dt.Columns.Add("Editorial", typeof(string));
                dt.Columns.Add("Año", typeof(int));

                // Variable para controlar si se encontraron resultados
                bool foundResults = false;

                // Iterar sobre las filas de Excel desde la fila 4 hasta la última
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Verificar si el valor en la columna de búsqueda coincide con el término de búsqueda
                    if (worksheet.Cells[row, searchColumn].Text.Contains(searchTerm))
                    {
                        // Si se encuentra, crear una nueva fila en el DataTable con los valores de Excel
                        DataRow newRow = dt.NewRow();
                        newRow["Título"] = worksheet.Cells[row, 1].Text;
                        newRow["Autor"] = worksheet.Cells[row, 2].Text;
                        newRow["ISBN"] = worksheet.Cells[row, 4].Text;
                        newRow["Editorial"] = worksheet.Cells[row, 5].Text;
                        newRow["Año"] = int.Parse(worksheet.Cells[row, 6].Text);

                        dt.Rows.Add(newRow); // Agregar la fila encontrada al DataTable
                        foundResults = true; // Marcar que se encontró un resultado
                    }
                }            

                // Asignar el DataTable al DataGridView para mostrar los resultados
                dataGridView1.DataSource = dt;
            }
        }

        private void eliminar_Click(object sender, EventArgs e)
        {
            retirar sexo = new retirar();
            sexo.MostrarPanel5();
            sexo.Show();
        }

        private void editar_Click(object sender, EventArgs e)
        {
            Modificar sexo = new Modificar();
            sexo.MostrarPanel7();
            sexo.Show();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo de búsqueda no esté vacío
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                LoadTesis(excelFilePath);
                return;
            }

            string searchTerm = textBox2.Text; // Lo que se buscará
            int searchColumn = 0; // Inicializa con la columna de búsqueda

            // Determinar cuál RadioButton está seleccionado para definir la columna de búsqueda
            if (radioButton9.Checked) searchColumn = 1; // Primera columna
            else if (radioButton8.Checked) searchColumn = 2; // Segunda columna
            else if (radioButton7.Checked) searchColumn = 4; // Tercera columna
            else if (radioButton6.Checked) searchColumn = 5; // Cuarta columna
            else if (radioButton1.Checked) searchColumn = 6; // Quinta columna

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Trabajamos con la primera hoja
                DataTable dt = new DataTable();

                // Crear columnas en el DataTable (Ajusta las columnas según tus necesidades)
                dt.Columns.Add("Título", typeof(string));
                dt.Columns.Add("Autor", typeof(string));
                dt.Columns.Add("Asesor", typeof(string));
                dt.Columns.Add("Carrera", typeof(string));
                dt.Columns.Add("Año", typeof(int));

                // Variable para controlar si se encontraron resultados
                bool foundResults = false;

                // Iterar sobre las filas de Excel desde la fila 4 hasta la última
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Verificar si el valor en la columna de búsqueda coincide con el término de búsqueda
                    if (worksheet.Cells[row, searchColumn].Text.Contains(searchTerm))
                    {
                        // Si se encuentra, crear una nueva fila en el DataTable con los valores de Excel
                        DataRow newRow = dt.NewRow();
                        newRow["Título"] = worksheet.Cells[row, 1].Text;
                        newRow["Autor"] = worksheet.Cells[row, 2].Text;
                        newRow["Asesor"] = worksheet.Cells[row, 4].Text;
                        newRow["Carrera"] = worksheet.Cells[row, 5].Text;
                        newRow["Año"] = int.Parse(worksheet.Cells[row, 6].Text);

                        dt.Rows.Add(newRow); // Agregar la fila encontrada al DataTable
                        foundResults = true; // Marcar que se encontró un resultado
                    }
                }

                // Asignar el DataTable al DataGridView para mostrar los resultados
                dataGridView2.DataSource = dt;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            agregar sexo = new agregar();
            sexo.MostrarPanelTesis();
            sexo.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            retirar sexo = new retirar();
            sexo.MostrarPanel7();
            sexo.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";
            LoadTesis(excelFilePath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Modificar sexo = new Modificar();
            sexo.MostrarPanel1();
            sexo.Show();
        }
    }
}


