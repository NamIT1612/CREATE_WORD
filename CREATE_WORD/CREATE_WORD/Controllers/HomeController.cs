using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.IO;
using Spire.Doc;

namespace CREATE_WORD.Controllers
{
    public class HomeController : Controller
    {
        // set up chữ thường
        public string setTextNormal(XWPFDocument doc, string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.LEFT;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r = p.CreateRun();
            r.SetText(str);
            r.FontFamily = "Time new roman";
            r.SetTextPosition(0);
            return str;
        }
        //set up chữ in
        public string setTextItalic(XWPFDocument doc, string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.DISTRIBUTE;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r = p.CreateRun();
            r.SetText(str);
            r.IsItalic = true;
            r.FontFamily = "Time new roman";
            r.SetTextPosition(0);
            return str;
        }
        //set up chữ nghiêng
        public string setTextBold(XWPFDocument doc ,string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.LEFT;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r = p.CreateRun();
            r.SetText(str);
            r.IsBold = true;
            r.FontFamily = "Time new roman";
            r.SetTextPosition(0);
            return str;
        }
        //set up chữ đậm và thường
        public string setTextBoldNormal(XWPFDocument doc, string str,string str1)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.LEFT;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r = p.CreateRun();
           
            XWPFRun r_1 = p.CreateRun();
            r.SetText(str);
            r_1.SetText(str1);
            r.IsBold = true;
            r.FontFamily = "Time new roman";
            r_1.FontFamily= "Time new roman";
            r.SetTextPosition(0);
            return str;
        }


        public string setTextNormalBold(XWPFDocument doc, string str1, string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.LEFT;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r_1 = p.CreateRun();
            XWPFRun r = p.CreateRun();
            r_1.SetText(str1);
            r.SetText(str);
            r.IsBold = true;
            r.FontFamily = "Time new roman";
            r_1.FontFamily = "Time new roman";
            r.SetTextPosition(0);
            return str;
        }
        public string setImgText(XWPFDocument doc, string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.VerticalAlignment = TextAlignment.CENTER;
            XWPFRun r1 = p.CreateRun();
            XWPFRun r = p.CreateRun();
            // Đường dẫn của file hình ảnh
            using (FileStream picFile = new FileStream(@"D:\\demo\\CREATE_WORD\\CREATE_WORD\\Img\\hv.png", FileMode.Open, FileAccess.Read))
            {
                r1.SetText("\t\t\t");
                r1.AddPicture(picFile, (int)PictureType.PNG, "hv", 200000, 200000);
            }
            
            r.SetText(str);
            r.FontSize = 11;
            r.FontFamily = "Time new roman";
            return str;
        }
        public string setTextNormalBoldItalic(XWPFDocument doc, string str1, string str)
        {
            XWPFParagraph p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.LEFT;
            p.VerticalAlignment = TextAlignment.TOP;
            XWPFRun r_1 = p.CreateRun();
            XWPFRun r = p.CreateRun();
            r_1.SetText(str1);
            r.SetText(str);
            r.IsBold = true;
            r.IsItalic = true;
            r_1.IsItalic = true;
            r.FontFamily = "Time new roman";
            r_1.FontFamily = "Time new roman";
            r.SetTextPosition(0);
            return str;
        }
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            XWPFDocument doc = new XWPFDocument();
            

            setTextNormalBoldItalic(doc, "Hôm nay, vào hồi ", "{ { Thời gian giao xe thực tế(giờ/ ngày / tháng / năm) } } , ");
            setTextNormal(doc, "");
            setTextNormal(doc, "");

            setTextBoldNormal(doc, "A.	Tình trạng xe khi nhận lại xe  ", "(đánh dấu tích để lựa chọn)");
                setTextBold(doc, "\t-	Tình trạng nội thất, ngoại thất và máy móc xe:");
                    setImgText(doc, "\tGiống tình trạng bạn đầu (nội thất, ngoại thất, máy móc, giấy tờ, đồ dự phòng)");
                    setImgText(doc, "\tKhác tình trạng ban đầu, các hư hỏng và mất mát sau:");
                    setTextBold(doc, "\t\t\t\t…………………………………………………………………………………………\n");
                    setTextBold(doc," \t\t\t\t Chi phí khắc phục(tạm tính) …………………..đ");


                 setTextBold(doc, "\t-	Số công tơ mét (Km): {{Số km nhận xe thực tế}}");
                 setTextNormalBold(doc, "\t\t Tổng số km đã đi:", "{ { Số km thực tế hành trình} }");
                    setImgText(doc, "\t Nằm trong giới hạn km");

                    XWPFParagraph p = doc.CreateParagraph();
                    p.VerticalAlignment = TextAlignment.CENTER;
                    XWPFRun pic = p.CreateRun();
                    XWPFRun r = p.CreateRun();
                    XWPFRun r1 = p.CreateRun();
                    XWPFRun r2 = p.CreateRun();
                    XWPFRun r3 = p.CreateRun();
                    using (FileStream picFile = new FileStream(@"D:\demo\CREATE_WORD\CREATE_WORD\Img\hv.png", FileMode.Open, FileAccess.Read))
                    {
                        pic.SetText("\t\t\t");
                        pic.AddPicture(picFile, (int)PictureType.PNG, "hv", 200000, 200000);
                    }
                    r.SetText(" Vượt số giới hạn km, số km vượt: ");
                    r.FontSize = 11;
                    r.FontFamily = "Time new roman";
                    r1.SetText("{{Số km phụ trội }} ");
                    r1.FontSize = 11;
                    r1.IsBold = true;
                    r1.FontFamily = "Time new roman";
                    r2.SetText("Số tiền phụ trội km: ");
                    r2.FontSize = 11;
                    r2.FontFamily = "Time new roman";
                    r3.SetText("{{Phí phụ trội km}}");
                    r3.FontSize = 11;
                    r3.IsBold = true;
                    r3.FontFamily = "Time new roman";
                    setTextBoldNormal(doc, "\t\t\t Đồng hồ xăng/dầu:", " (vạch xăng):");
                    setTextBold(doc, " \t\t\t Phụ phí xăng dầu ");
            setTextBold(doc, "\t-	Thời gian phụ trội so với hợp đồng: {{Thời gian trả xe quá }} giờ, phụ phí phát sinh {{Phí giao muộn}}");
            setTextBold(doc, "\t-	Vé cầu đường phát sinh chưa thanh toán: {{Phí phụ trội ETC}}");
            setTextBold(doc, "\t-	Các lỗi ghi nhận được trong quá trình thuê:	");
                    setImgText(doc, "\tChưa phát hiện lỗi gì");
                    setImgText(doc, "\tPhát hiện lỗi:");
                    setTextNormal(doc, "\t\t\t\t\t-\tVượt tốc độ");
                    setTextNormal(doc, "\t\t\t\t\t-\tVào đường cấm:  ……………………………");
                    setTextNormal(doc, "\t\t\t\t\t-\tVượt đèn đỏ: ………………………………...");
                    setTextNormal(doc, "\t\t\t\t\t-\tCác lỗi khác(nếu có): ……………..");

            setTextBold(doc, "Tổng chi phí phát sinh so với hợp đồng: {{Tổng phát sinh thêm}}");
            setTextBold(doc, "Bên A đã hoàn trả cho bên B một số giấy tờ và tài sản như sau:");
                    setImgText(doc, "\tToàn bộ giấy tờ và tài sản tại thời điểm giao nhận");
                    setImgText(doc, "\tThiếu hoặc chưa hoàn trả các giấy tờ và tài sản sau: ");
                    setTextItalic(doc, "\t\t\t\t………………………………………………………………………………………");
                    setTextItalic(doc, "\t\t\t\t………………………………………………………………………………………");
            setTextBold(doc, "Cam kết của Khách thuê và Chủ xe:	");

            setTextItalic(doc, "\tTrong trường hợp phát sinh các khoản phạt nguội, bằng chứng do camera giám sát của Cục CSGT - Bộ Công An ghi nhận được trong thời gian Bên B sử dụng xe ô tô thuê của Bên A. Bên A có trách nhiệm cung cấp các bằng chứng liên quan cho bên B ngay khi nhận được thông tin. Bên B cam kết chịu hoàn toàn trách nhiệm và bồi thường toàn bộ các chi phí liên quan cho Bên A. ");

            XWPFTable table2 = doc.CreateTable(2, 2);
            table2.Width = 5000;
            table2.GetRow(0).GetCell(0).SetText("\t\t\t\tBên A");
            table2.GetRow(0).GetCell(1).SetText("\t\t\t\tBên B");
            table2.SetColumnWidth(0, 2000);


            // liên kết đường dẫn và tạo file word
            // Lưu ý dường dẫn
            using (
                FileStream fs = new FileStream("D:\\LTDD\\complexTable.docx", FileMode.Create))
            {
                doc.Write(fs);
            }
            return View();
        }
    }
}
