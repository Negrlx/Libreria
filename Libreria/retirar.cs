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
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace Libreria
{
    public partial class retirar : Form
    {        
        public retirar()
        {
            InitializeComponent();
        }

        private Base_Estudiante baseEstudiante;

        public retirar(Base_Estudiante formA)
        {
            InitializeComponent();
            baseEstudiante = formA; 
        }

        private void SearchAndUpdateQuantityInExcel()
        {
            int cont = CountRowsFromExcel();
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";

            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("El campo del código no puede estar vacío.");
                return;
            }

            string codeToSearch = textBox1.Text;
            bool codeFound = false;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Primera hoja

                for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 4].Text == codeToSearch) // Busca en la columna 4
                    {
                        string isbn = worksheet.Cells[row, 4].Text;
                        string title = worksheet.Cells[row, 1].Text;
                        string author = worksheet.Cells[row, 2].Text;
                        string year = worksheet.Cells[row, 6].Text;
                        string quantity = worksheet.Cells[row, 3].Text;
                        string editorial = worksheet.Cells[row, 5].Text;

                        int currentQuantity;
                        if (int.TryParse(quantity, out currentQuantity))
                        {
                            if (currentQuantity > 0)
                            {
                                currentQuantity--; // Restar uno a la cantidad de ejemplares
                                worksheet.Cells[row, 3].Value = currentQuantity; // Actualizar la cantidad en la celda

                                // Mostrar un mensaje con los datos y la cantidad actualizada
                                MessageBox.Show($"Código encontrado:\nISBN: {isbn}\nTítulo: {title}\nAutor: {author}\nAño: {year}\nCantidad actualizada: {currentQuantity}\nEditorial: {editorial}");

                                codeFound = true;

                                // Guardar los cambios en el archivo Excel (Primera hoja)
                                package.Save();

                                

                                // Ahora agregar el ISBN a la tercera hoja si no existe
                                AddISBNToHistory(isbn, package);

                            }
                            else
                            {
                                MessageBox.Show($"No hay ejemplares disponibles para el código: {isbn}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("El valor de la cantidad es más de lo disponible.");
                        }

                        break;
                    }
                }

                if (!codeFound)
                {
                    MessageBox.Show("El código no fue encontrado en la tabla de Excel.");
                }
            }
        }

        private void AddISBNToHistory(string isbn, ExcelPackage package)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[2]; // Tercera hoja

            bool isbnExists = false;
            for (int rowt = 2; rowt <= worksheet.Dimension.End.Row; rowt++) // Lee hasta la última fila
            {
                string cellISBN = NormalizeISBN(worksheet.Cells[rowt, 1].Text);

                // Comparar ISBN
                if (cellISBN == NormalizeISBN(isbn))
                {
                    isbnExists = true;
                    break; // Salir si el ISBN ya existe
                }
            }

            // Si no existe el ISBN, agregar una nueva fila
            if (!isbnExists)
            {
                int newRow = worksheet.Dimension.End.Row + 1;
                worksheet.Cells[newRow, 1].Value = isbn;

                // Guardar cambios en la hoja 3
                package.Save();

                MessageBox.Show($"ISBN {isbn} registrado en el historial de retiros.");
            }
        }
       
        private string NormalizeISBN(string isbn)
        {
            return isbn.Replace("-", "").Trim().ToLower();
        }           

        private void button1_Click(object sender, EventArgs e)
        {
            
                int cont = CountRowsFromExcel();

                if (cont < 3)
                {
                    SearchAndUpdateQuantityInExcel();
                }
                else
                {
                    MessageBox.Show("Ya has retirado la cantidad máxima de libros disponible.\nRegresa los que debes.\n\nMoroso.");
                }
            
        }

        public void MostarPanel1()
        {
            panel1.Show(); // Mostrar Panel1
            panel3.Hide();
            panel5.Hide();
            panel7.Hide();

        }

        public void MostrarPanel3()
        {
            panel3.Show(); // Mostrar Panel3
            panel1.Hide();
            panel5.Hide();
            panel7.Hide();

        }

        public void MostrarPanel5()
        {
            panel3.Hide(); // Mostrar Panel5
            panel1.Hide();
            panel5.Show();
            panel7.Hide();

        }

        public void MostrarPanel7()
        {
            panel3.Hide(); // Mostrar Panel7
            panel1.Hide();
            panel5.Hide();
            panel7.Show();
        }

        private void ReturnBook()
        {
            List<string> isbnRetirados = GetIsbnRetiradosFromExcel(); // Obtener la lista de libros retirados
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Ruta al archivo de Excel

            if (string.IsNullOrWhiteSpace(textBox2.Text)) // Verificar que el campo de texto no esté vacío
            {
                MessageBox.Show("El campo del código no puede estar vacío.");
                return;
            }

            string codeToReturn = textBox2.Text; // Código del libro que se va a devolver

            if (isbnRetirados.Count == 0) // Verificar si hay libros retirados
            {
                MessageBox.Show("No hay libros retirados para devolver.");
                return;
            }

            if (!isbnRetirados.Contains(codeToReturn)) // Verificar si el ISBN está en la lista de retirados
            {
                MessageBox.Show("El código ingresado no está en la lista de libros retirados.");
                return;
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet sheetHistorial = package.Workbook.Worksheets[2]; // Hoja 3: Historial de retiros
                ExcelWorksheet sheetLibros = package.Workbook.Worksheets[0];    // Hoja 1: Libros

                bool codeFound = false; // Variable para saber si se encuentra el código

                // Paso 1: Eliminar el ISBN de la hoja de historial (hoja 3)
                for (int row = 2; row <= sheetHistorial.Dimension.End.Row; row++)
                {
                    if (sheetHistorial.Cells[row, 1].Text == codeToReturn) // Verificar el ISBN en la columna 1
                    {
                        // Eliminar la fila completa
                        sheetHistorial.DeleteRow(row);
                        package.Save(); // Guardar después de eliminar la fila del historial
                        codeFound = true;
                        break; // Salir del bucle una vez encontrado y eliminado el ISBN
                    }
                }

                if (!codeFound) // Si no se encuentra el código en el historial
                {
                    MessageBox.Show("El código no fue encontrado en el historial de retiros.");
                    return;
                }

                // Paso 2: Actualizar la cantidad en la hoja de libros (hoja 1)
                codeFound = false; // Reiniciar la variable para la búsqueda en la hoja de libros
                for (int row = 2; row <= sheetLibros.Dimension.End.Row; row++)
                {
                    if (sheetLibros.Cells[row, 4].Text == codeToReturn) // Verificar el ISBN en la columna 1
                    {
                        string quantity = sheetLibros.Cells[row, 3].Text; // Columna 3: Cantidad
                        int currentQuantity;

                        if (int.TryParse(quantity, out currentQuantity))
                        {
                            currentQuantity++; // Incrementar la cantidad en 1
                            sheetLibros.Cells[row, 3].Value = currentQuantity; // Actualizar la celda con la nueva cantidad
                            package.Save(); // Guardar los cambios en el archivo Excel
                            codeFound = true;

                            // Mostrar mensaje de éxito
                            MessageBox.Show($"El libro con ISBN {codeToReturn} ha sido devuelto correctamente.\nCantidad actualizada: {currentQuantity}");

                            break; // Salir del bucle una vez actualizado
                        }
                        else
                        {
                            MessageBox.Show("Error al leer la cantidad del libro.");
                        }
                    }
                }

                if (!codeFound) // Si no se encuentra el código en la tabla de libros
                {
                    MessageBox.Show("El código no fue encontrado en la tabla de libros.");
                }
            }
        }

        private List<string> GetIsbnRetiradosFromExcel()
        {
            List<string> isbnRetirados = new List<string>(); // Lista para almacenar los ISBN retirados
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Ruta del archivo Excel

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2]; // Trabajar con la tercera hoja (historial de retiros)

                // Verificar que la hoja de cálculo no esté vacía
                if (worksheet.Dimension == null)
                {
                    MessageBox.Show("La hoja de historial está vacía o no tiene datos.");
                    return isbnRetirados; // Devolver la lista vacía si no hay datos
                }

                // Iterar sobre las filas que contienen los ISBN retirados (asumiendo que comienzan en la fila 2)
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string isbn = worksheet.Cells[row, 1].Text; // Columna 1: ISBN retirado

                    if (!string.IsNullOrWhiteSpace(isbn)) // Asegurarse de que el ISBN no esté vacío
                    {
                        isbnRetirados.Add(isbn); // Agregar el ISBN a la lista
                    }
                }
            }

            return isbnRetirados; // Devolver la lista de ISBN retirados
        }

        private int CountRowsFromExcel()
        {
            try
            {
                string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx";

                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[2]; // Tercera hoja

                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show("No hay filas en la hoja de Excel.");
                        return 0;
                    }

                    // Contar las filas existentes desde la fila 2 hasta la última fila
                    int rowCount = worksheet.Dimension.End.Row; // -1 para no contar la fila de encabezado

                    return rowCount; // Retornar el número de filas sin modificar el archivo Excel
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error al procesar el archivo Excel: {ex.Message}");
                return 0;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int cont = CountRowsFromExcel();

            if (cont > 0 )
            {
                ReturnBook();
            }
            else
            {
                MessageBox.Show("No tienes los cuales puedas devolver");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo del código no esté vacío
            if (string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("El campo del código no puede estar vacío.");
                return;
            }

            string codeToSearch = textBox3.Text; // Código que se buscará
            bool codeFound = false; // Variable para saber si el código se encontró

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Trabajar con la primera hoja

                // Iterar sobre las filas desde la fila 4 hasta la última
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Buscar en la columna 4 el código
                    if (worksheet.Cells[row, 4].Text == codeToSearch)
                    {
                        // Si se encuentra el código, eliminar la fila
                        worksheet.DeleteRow(row);

                        // Marcar que se encontró el código
                        codeFound = true;

                        // Guardar los cambios en el archivo
                        FileInfo file = new FileInfo(excelFilePath);
                        package.SaveAs(file);

                        // Mostrar mensaje de confirmación
                        MessageBox.Show($"La fila con el código {codeToSearch} ha sido eliminada correctamente.");

                        break; // Salir del bucle si se encontró el código
                    }
                }

                // Si no se encontró el código, mostrar un mensaje de error
                if (!codeFound)
                {
                    MessageBox.Show("El código no fue encontrado en la tabla de Excel.");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"C:\Users\Santiago\Desktop\Libro1.xlsx"; // Cambia esto a la ruta de tu archivo Excel

            // Validar que el campo del código no esté vacío
            if (string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("El campo del código no puede estar vacío.");
                return;
            }

            string codeToSearch = textBox4.Text; // Código que se buscará
            bool codeFound = false; // Variable para saber si el código se encontró

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Trabajar con la primera hoja

                // Iterar sobre las filas desde la fila 4 hasta la última
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Buscar en la columna 4 el código
                    if (worksheet.Cells[row, 1].Text == codeToSearch)
                    {
                        // Si se encuentra el código, eliminar la fila
                        worksheet.DeleteRow(row);

                        // Marcar que se encontró el código
                        codeFound = true;

                        // Guardar los cambios en el archivo
                        FileInfo file = new FileInfo(excelFilePath);
                        package.SaveAs(file);

                        // Mostrar mensaje de confirmación
                        MessageBox.Show($"La fila con el Tirulo de la Tesis {codeToSearch} ha sido eliminada correctamente.");

                        break; // Salir del bucle si se encontró el código
                    }
                }

                // Si no se encontró el código, mostrar un mensaje de error
                if (!codeFound)
                {
                    MessageBox.Show("El Tirulo de la Tesis no fue encontrado en la tabla de Excel.");
                }
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
