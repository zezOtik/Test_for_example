using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Sockets;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Diagnostics;
using System.Runtime.InteropServices;

using MessagingToolkit.QRCode.Codec;/*Библиотека для QR code*/
using MessagingToolkit.QRCode.Codec.Data;

using itextsharp;/*Библиотека для работы с пдф документами*/
using iTextSharp.text;
using iTextSharp.text.pdf;

using System.Net;//для отправки сообщений
using System.Net.Mail;


namespace Test_for_example
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }

        private DataTable ReadCSVFile(string pathToCsvFile)
        {
            /*Данные об аккаунте через который будут отправляться сообщения*/

            SmtpClient client = new SmtpClient("smtp.mail.ru", 587);
            client.Credentials = new NetworkCredential("zotik_65@mail.ru", "mzprophylaxis02");
            client.EnableSsl = true;
            string from = "zotik_65@mail.ru";

            DataTable table = new DataTable("Customers");//Новая таблица с названием Клиенты

   
            DataColumn columnEmail;
            columnEmail = new DataColumn("Email", typeof(String));//Колонка для мыл

            DataColumn columnName;
            columnName = new DataColumn("Name", typeof(String));//Колонка для имен

            DataColumn columnCompany;
            columnCompany = new DataColumn("Company", typeof(String));//Колонка Компания

            DataColumn columnYear;
            columnYear = new DataColumn("Year", typeof(Int32));//Колонка возраст

            List<string> emailsList = new List<string>();//Лист для мыл, для работы без залазки обратно в файл, чтобы отправить письма

            table.Columns.AddRange(new DataColumn[] {columnEmail, columnName, columnCompany, columnYear});//Добавление таблицы
            
            try
            {
                DataRow row;//Новая строка

                string[] customersValue;//Строка

                string[] customers = File.ReadAllLines(pathToCsvFile);//Считать весь файл

                for (int i = 0; i < customers.Length; ++i)//Цикл записи
                {
                    if (!String.IsNullOrEmpty(customers[i]))//Пока не пуст читать
                    {
                        customersValue = customers[i].Split(',');//Объявление разделителя

                       // emailsList.Add(customersValue[0]);//Добавили мыло в лист
                        
                        row = table.NewRow();
                       
                        row["Email"] = customersValue[0];//считывание данных из файлы
                        row["Name"] = customersValue[1];
                        row["Company"] = customersValue[2];
                        row["Year"] = int.Parse(customersValue[3]);

                        string QRText = (customersValue[0] + customersValue[1] + customersValue[2] + int.Parse(customersValue[3]));//Генерация строки для QR code

                        QRCodeEncoder QREncoder = new QRCodeEncoder(); //Объявление переменной
                        Bitmap QRCode = QREncoder.Encode(QRText);//Битмап для кьюр картинки

                        QRCode.Save(Application.StartupPath + @"\image.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);//Сохранение в картинку

                        Document document = new Document();//Переменная документ
                        
                        PdfWriter.GetInstance(document, new FileStream(Application.StartupPath + String.Format(@"\QR_1.pdf"), FileMode.Create));//В потоке отправляем картинку и запиливаем в пдф

                        document.Open();//Открытие 
                        
                        iTextSharp.text.Image PDFImage = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\image.jpeg");//Добавление картинки в пдф
                        PDFImage.Alignment = Element.ALIGN_CENTER;

                        document.Add(PDFImage);

                        PdfPTable PDFTable = new PdfPTable(3);//Добавление таблички в пдф с фразкой

                        PdfPCell PDFCell = new PdfPCell(new Phrase("Hello," + customersValue[1], new iTextSharp.text.Font(iTextSharp.text.Font.TIMES_ROMAN, 20, iTextSharp.text.Font.NORMAL)));
                        PDFCell.Padding = 5;
                        PDFCell.Colspan = 3;
                        PDFCell.HorizontalAlignment = Element.ALIGN_CENTER;

                        PDFTable.AddCell(PDFCell);

                        document.Add(PDFTable);


                        document.Close();//Закрывание пдф файла

                        File.Delete(Application.StartupPath + @"\image.jpeg");//удаляем картинку, чтобы не засорять
                        
                        table.Rows.Add(row);//Новая строка

                        string to = customersValue[0];
                        string subject = "Hello";
                        string text = "Hello" + customersValue[1];
                        MailMessage message = new MailMessage(from, to, subject, text);
                       // StreamReader stream = new StreamReader();
                        Attachment sendfile = new Attachment(@"C:\Users\zotik\Documents\Visual Studio 2013\Projects\dz_ayp_2kyrs\Test_for_example\Test_for_example\bin\Debug\QR_1.pdf");
                        message.Attachments.Add(sendfile);
                        client.Send(message);
                        message.Attachments.Dispose();//Не дает обратиться к файлу, если он используется

                    }

                    File.Delete(Application.StartupPath + @"\QR_1.pdf");//удаляем картинку, чтобы не засорять

                    
                    
                }
                MessageBox.Show("Все отправлено");
              
            }
           
            catch (Exception exception)//Ошибка о неправильном считывание
            {
                MessageBox.Show(exception.Message);
            }

            return table;
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = ReadCSVFile(@"C:\Users\zotik\Documents\Visual Studio 2013\Projects\dz_ayp_2kyrs\Test_for_example\files\customers.csv");//Путь к цсв файлу
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
   
    }
 }
   

