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

namespace Libreria
{
    public partial class agregar : Form
    {
        public agregar()
        {
            InitializeComponent();
        }

        private void AddBookToExcel()
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el ISBN no esté vacío
            if (new[] { txtISBN.Text }.Any(string.IsNullOrWhiteSpace))
            {
                MessageBox.Show("No puede haber espacios vacios");
                return;
            }

            // Variable para verificar si el ISBN existe
            bool isbnExists = false;

            

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Asume que estás trabajando con la primera hoja

                // Iterar sobre las filas para buscar el ISBN

                // Iterar sobre las filas para buscar el ISBN
                for (int row = 4; row <= 50; row++) // Lee hasta la fila 50
                {
                    // Normalizar el ISBN de la celda y el ISBN ingresado
                    string cellISBN = NormalizeISBN(worksheet.Cells[row, 1].Text);
                    string inputISBN = NormalizeISBN(txtISBN.Text);

                    // Mostrar los valores que se están comparando (esto es solo para depuración)
                    Console.WriteLine($"Comparando: '{cellISBN}' con '{inputISBN}'");

                    // Comparar los ISBN
                    if (cellISBN == inputISBN)
                    {
                        int existingQuantity;
                        // Verifica que la cantidad existente pueda ser analizada correctamente
                        if (int.TryParse(worksheet.Cells[row, 5].Text, out existingQuantity)) // Cambia 5 por el índice de columna de Cantidad
                        {
                            existingQuantity += int.Parse(txtQuantity.Text); // Sumar la cantidad
                            worksheet.Cells[row, 5].Value = existingQuantity; // Actualizar la cantidad en la celda
                            isbnExists = true;
                            break; // Salir del bucle si se encontró el ISBN
                        }
                    }
                }


                // Si el ISBN no existe, agregar una nueva fila
                if (!isbnExists)
                {
                    int newRow = worksheet.Dimension.End.Row + 1; // Encuentra la próxima fila vacía
                    worksheet.Cells[newRow, 4].Value = txtISBN.Text; // ISBN
                    worksheet.Cells[newRow, 1].Value = txtTitle.Text; // Título
                    worksheet.Cells[newRow, 2].Value = txtAuthor.Text; // Autor
                    worksheet.Cells[newRow, 6].Value = txtYear.Text; // Año
                    worksheet.Cells[newRow, 3].Value = txtQuantity.Text; // Cantidad
                    worksheet.Cells[newRow, 5].Value = txteditorial.Text; //Editorial
                }

                // Guardar los cambios en el archivo Excel
                package.Save();
            }

            // Verificar si el archivo fue actualizado correctamente
            if (File.Exists(excelFilePath))
            {
                using (ExcelPackage checkPackage = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet checkWorksheet = checkPackage.Workbook.Worksheets[0];
                    bool updatedSuccessfully = false;

                    // Revisar si el último ISBN añadido o actualizado es el correcto
                    for (int row = 4; row <= checkWorksheet.Dimension.End.Row; row++)
                    {
                        if (checkWorksheet.Cells[row, 1].Text == txtISBN.Text) // Cambia 1 por el índice de columna de ISBN
                        {
                            updatedSuccessfully = true;
                            break;
                        }
                    }

                    // Mensaje de confirmación
                    if (updatedSuccessfully)
                    {
                        MessageBox.Show("La tabla de Excel se actualizó correctamente.");
                    }
                    else
                    {
                        MessageBox.Show("Hubo un problema al actualizar la tabla de Excel.");
                    }
                }
            }
            else
            {
                MessageBox.Show("El archivo de Excel no existe.");
            }

            // Limpiar los TextBox después de agregar la fila
            txtISBN.Clear();
            txtTitle.Clear();
            txtAuthor.Clear();
            txtYear.Clear();
            txtQuantity.Clear();
        }

        private string NormalizeISBN(string isbn)
        {
            return isbn.Replace("-", "").Trim().ToLower();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddBookToExcel();
            this.Close();
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void AddThesisToExcel()
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Asume que estás trabajando con la primera hoja


                int newRow = worksheet.Dimension.End.Row + 1; // Encuentra la próxima fila vacía
                worksheet.Cells[newRow, 1].Value = textBox12.Text; // Titulo
                worksheet.Cells[newRow, 2].Value = textBox8.Text; // Autor
                worksheet.Cells[newRow, 3].Value = textBox9.Text; // Asesor
                worksheet.Cells[newRow, 4].Value = textBox10.Text; // Carrera
                worksheet.Cells[newRow, 5].Value = textBox11.Text; // Año

                // Guardar los cambios en el archivo Excel
                package.Save();
            }

            // Verificar si el archivo fue actualizado correctamente
            if (File.Exists(excelFilePath))
            {
                using (ExcelPackage checkPackage = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet checkWorksheet = checkPackage.Workbook.Worksheets[1];
                    bool updatedSuccessfully = false;

                    // Revisar si el último ISBN añadido o actualizado es el correcto
                    for (int row = 4; row <= checkWorksheet.Dimension.End.Row; row++)
                    {
                        if (checkWorksheet.Cells[row, 1].Text == txtISBN.Text) // Cambia 1 por el índice de columna de ISBN
                        {
                            updatedSuccessfully = true;
                            break;
                        }
                    }

                    // Mensaje de confirmación
                    if (updatedSuccessfully)
                    {
                        MessageBox.Show("La tabla de Excel se actualizó correctamente.");
                    }
                    else
                    {
                        MessageBox.Show("Hubo un problema al actualizar la tabla de Excel.");
                    }
                }
            }
            else
            {
                MessageBox.Show("El archivo de Excel no existe.");
            }

            // Limpiar los TextBox después de agregar la fila
            textBox10.Clear();
            textBox11.Clear();
            textBox12.Clear();
            textBox8.Clear();
            textBox9.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            AddThesisToExcel();
            this.Close();
        }

        public void MostrarPanelTesis()
        {
            panel15.Show(); // Mostrar Panel3
            panel1.Hide();
        }

        public void MostrarPanelLibro()
        {
            panel15.Hide(); // Mostrar Panel3
            panel1.Show();
        }

        private void agregar_Load(object sender, EventArgs e)
        {

        }
    }
}
