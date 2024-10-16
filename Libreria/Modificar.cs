using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Libreria
{
    public partial class Modificar : Form
    {
        public Modificar()
        {
            InitializeComponent();
        }

        // Variable global para almacenar la fila encontrada
        private int foundRow = -1;

        private void button4_Click(object sender, EventArgs e)
        {
            entradal();
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            foundRow = -1; // Reiniciar la fila encontrada

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja

                // Iterar sobre las filas desde la fila 4 hasta la última
                for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Buscar en la columna 4 el código (ISBN)
                    if (worksheet.Cells[row, 4].Text == textBox4.Text)
                    {
                        foundRow = row; // Guardar la fila donde se encontró el código
                        MessageBox.Show("Código encontrado en la fila: " + row); // Mostrar la fila donde se encontró
                        break; // Salir del bucle si se encontró el código
                    }
                }
            }

            if (foundRow == -1)
            {
                MessageBox.Show("El código no fue encontrado en la tabla de Excel.");
            }
        }

        private void ModifyRowInExcel(int row)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajamos con la primera hoja

                // Si la fila es válida (es decir, el código fue encontrado), modificar la fila
                if (row > 0)
                {
                    // Solo se actualizan las celdas si los TextBox tienen información
                    if (!string.IsNullOrWhiteSpace(textBox12.Text)) worksheet.Cells[row, 1].Value = textBox12.Text; // Columna 1: Título
                    if (!string.IsNullOrWhiteSpace(textBox8.Text)) worksheet.Cells[row, 2].Value = textBox8.Text; // Columna 2: Autor
                    if (!string.IsNullOrWhiteSpace(textBox11.Text)) worksheet.Cells[row, 3].Value = textBox11.Text; // Columna 3: Cantidad
                    if (!string.IsNullOrWhiteSpace(textBox9.Text)) worksheet.Cells[row, 4].Value = textBox9.Text; // Columna 4: ISBN
                    if (!string.IsNullOrWhiteSpace(textBox10.Text)) worksheet.Cells[row, 5].Value = textBox10.Text; // Columna 5: Editorial
                    if (!string.IsNullOrWhiteSpace(textBox7.Text)) worksheet.Cells[row, 6].Value = textBox7.Text; // Columna 6: Año

                    // Guardar los cambios en el archivo Excel
                    FileInfo file = new FileInfo(excelFilePath);
                    package.SaveAs(file);

                    // Mostrar mensaje de confirmación
                    MessageBox.Show("La fila ha sido modificada correctamente.");
                }
                else
                {
                    MessageBox.Show("No se encontró la fila para modificar.");
                }
            }
        }

        private void ModifyRowInExcelThesis(int row)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Trabajamos con la primera hoja

                // Si la fila es válida (es decir, el código fue encontrado), modificar la fila
                if (row > 0)
                {
                    // Solo se actualizan las celdas si los TextBox tienen información
                    if (!string.IsNullOrWhiteSpace(textBox14.Text)) worksheet.Cells[row, 1].Value = textBox14.Text; // Columna 1: Título
                    if (!string.IsNullOrWhiteSpace(textBox3.Text)) worksheet.Cells[row, 2].Value = textBox3.Text; // Columna 2: Autor
                    if (!string.IsNullOrWhiteSpace(textBox5.Text)) worksheet.Cells[row, 3].Value = textBox5.Text; // Columna 3: Asesor
                    if (!string.IsNullOrWhiteSpace(textBox6.Text)) worksheet.Cells[row, 4].Value = textBox6.Text; // Columna 4: Carrera
                    if (!string.IsNullOrWhiteSpace(textBox13.Text)) worksheet.Cells[row, 5].Value = textBox13.Text; // Columna 5: Año

                    // Guardar los cambios en el archivo Excel
                    FileInfo file = new FileInfo(excelFilePath);
                    package.SaveAs(file);

                    // Mostrar mensaje de confirmación
                    MessageBox.Show("La fila ha sido modificada correctamente.");
                }
                else
                {
                    MessageBox.Show("No se encontró la fila para modificar.");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (foundRow > 0)
            {
                ModifyRowInExcel(foundRow); // Modificar la fila con los datos de los TextBox
            }
            else
            {
                MessageBox.Show("Primero debes buscar el código antes de modificar.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (foundRow > 0)
            {
                ModifyRowInExcelThesis(foundRow); // Modificar la fila con los datos de los TextBox
            }
            else
            {
                MessageBox.Show("Primero debes buscar el código antes de modificar.");
            }
        }

        public void MostrarPanel7()
        {
            // Asegúrate de que el nombre del panel sea correcto
            this.MaximumSize = new Size(342, 204);
            this.MinimumSize = new Size(342, 204);
            this.Size = new Size(342, 204);
            panel15.Hide();
            panel7.Show();
            panel3.Hide();
            panel1.Hide();
        }

        public void MostrarPanel1()
        {
            // Asegúrate de que el nombre del panel sea correcto
            this.MaximumSize = new Size(342, 204);
            this.MinimumSize = new Size(342, 204);
            this.Size = new Size(342, 204);
            panel15.Hide();
            panel7.Hide();
            panel3.Hide();
            panel1.Show();
        }

        public void entradal()
        {
            // Asegúrate de que el nombre del panel sea correcto
            this.MaximumSize = new Size(408, 545);
            this.MinimumSize = new Size(408, 545);
            this.Size = new Size(408, 545);
            panel15.Show();
            panel7.Hide();
            panel3.Hide();
            panel1.Hide();
        }

        public void entradat()
        {
            // Asegúrate de que el nombre del panel sea correcto
            this.MaximumSize = new Size(408, 545);
            this.MinimumSize = new Size(408, 545);
            this.Size = new Size(408, 545);
            panel15.Hide();
            panel7.Hide();
            panel3.Show();
            panel1.Hide();
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            entradat();
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            foundRow = -1; // Reiniciar la fila encontrada

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Trabajamos con la primera hoja

                // Iterar sobre las filas desde la fila 4 hasta la última
                for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Buscar en la columna 4 el código (ISBN)
                    if (worksheet.Cells[row, 1].Text == textBox1.Text)
                    {
                        foundRow = row; // Guardar la fila donde se encontró el código
                        MessageBox.Show("Código encontrado en la fila: " + row); // Mostrar la fila donde se encontró
                        break; // Salir del bucle si se encontró el código
                    }
                }
            }

            if (foundRow == -1)
            {
                MessageBox.Show("El código no fue encontrado en la tabla de Excel.");
            }
        }

        private void Modificar_Load(object sender, EventArgs e)
        {

        }
    }
}
