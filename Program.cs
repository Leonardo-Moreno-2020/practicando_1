using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SpreadsheetLight;

namespace practicando_1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string path = File.ReadAllText(@"C:\Users\lmorenom\Documents\DATOS.txt");
                string destino = @"C:\Users\lmorenom\Documents\Libros de ejercicios excel\Libro_pruebas.xlsx";

                Console.WriteLine("Los datos del elemento:");
                Console.WriteLine(path);

                string[] elementos = path.Replace(Environment.NewLine, ":").Split(':'); // division del arreglo
                Console.WriteLine("El texto ha sido divido");
                Console.WriteLine(elementos.Length);

                using (SLDocument doc = new SLDocument())
                {
                    // ENCABEZADO
                    doc.SetCellValue("A1", "Nombre");
                    doc.SetCellValue("B1", "Cedula");
                    doc.SetCellValue("C1", "Celular");
                    doc.SetCellValue("D1", "Correo");

                    // FILAS Y COLUMNAS DEL ARREGLO
                    doc.SetCellValue("A2", elementos[0]);
                    doc.SetCellValue("B2", elementos[1]);
                    doc.SetCellValue("C2", elementos[2]);
                    doc.SetCellValue("D2", elementos[3]);
                    doc.SetCellValue("A3", elementos[5]); //cambio de fila nombre
                    doc.SetCellValue("B3", elementos[6]);
                    doc.SetCellValue("C3", elementos[7]);
                    doc.SetCellValue("D3", elementos[8]);
                    doc.SetCellValue("A4", elementos[10]); //cambio de fila nombre
                    doc.SetCellValue("B4", elementos[11]);
                    doc.SetCellValue("C4", elementos[12]);
                    doc.SetCellValue("D4", elementos[13]);
                    doc.SetCellValue("A5", elementos[15]);
                    doc.SetCellValue("B5", elementos[16]);
                    doc.SetCellValue("C5", elementos[17]);
                    doc.SetCellValue("D5", elementos[18]);

                    doc.SaveAs(destino);

                    if (doc != null)
                    {
                        string ruta = @"C:\Users\lmorenom\Documents\Libros de ejercicios excel\mensaje.txt";
                        string mensaje = "Archivos y datos creados correctamente";

                        using (StreamWriter mensaje1 = new StreamWriter(ruta))
                        {
                            mensaje1.Write(mensaje);
                        }
                    }
                    else
                    {
                        string fallo = @"C:\Users\lmorenom\Documents\Libros de ejercicios excel\error.txt";
                        string mensaje2 = "Ha ocurrido un error en la operacion solicitada";

                        using (StreamWriter ST = new StreamWriter(fallo))
                        {
                            ST.Write(mensaje2);
                        }
                    }
                }

            }
            catch (Exception except)
            {
                Console.WriteLine(except.Message);
            }
      
                Console.ReadKey();
        }
    }
}
