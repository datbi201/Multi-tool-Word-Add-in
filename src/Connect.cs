// Trợ lý Kỹ thuật - Rev0.4: điểm vào của COM add-in cho Word (không cần Visual Studio, không cần PIA).
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace TroLyKyThuat
{
    public enum ext_ConnectMode { ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2, ext_cm_CommandLine = 3, ext_cm_Solution = 4, ext_cm_UISetup = 5 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1, ext_dm_UISetupComplete = 2, ext_dm_SolutionClosed = 3 }

    // Khai báo giống hệt bản PIA của Microsoft (Extensibility.dll): giao diện kép (dual), có DispId.
    // Cùng GUID và thứ tự hàm như IDTExtensibility2 của Microsoft (giao diện kép: 7 hàm IDispatch rồi 5 hàm này).
    // Mọi tham số mảng/đối tượng khai báo là con trỏ thô (IntPtr) để CLR không phải tạo mã chuyển đổi kiểu,
    // vì lỗi ở lớp chuyển đổi này làm sập clr.dll (mã 0x80131506) trên máy bạn.
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType((short)4160)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection(IntPtr Application, int ConnectMode, IntPtr AddInInst, IntPtr custom);
        [DispId(2)] void OnDisconnection(int RemoveMode, IntPtr custom);
        [DispId(3)] void OnAddInsUpdate(IntPtr custom);
        [DispId(4)] void OnStartupComplete(IntPtr custom);
        [DispId(5)] void OnBeginShutdown(IntPtr custom);
    }

    // Giống bản PIA của Office (office.dll): giao diện kép, DispId 1.
    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), TypeLibType((short)4160)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    /// <summary>
    /// Cầu nối với luồng chính của Word. Mọi lệnh gọi Word đều phải chạy trên luồng này (nhanh hơn nhiều so với gọi chéo luồng,
    /// và Word không bị kẹt hộp thoại). Bảng điều khiển chạy trên luồng riêng nên phải gọi Word qua OnWord.
    /// </summary>
    public static class Host
    {
        public static object App;
        static Control sync;
        static int mainId;

        public static void Init(object app)
        {
            App = app;
            try
            {
                mainId = Thread.CurrentThread.ManagedThreadId;
                sync = new Control();
                IntPtr h = sync.Handle;      // tạo handle ngay trên luồng chính của Word
                Log.W("Host: da tao cau noi luong chinh, thread " + mainId);
                StartPoll();
            }
            catch (Exception ex) { sync = null; Log.W("Host.Init loi: " + ex); }
        }

        // Bộ đếm giờ trên luồng chính của Word: theo dõi vùng chọn để gom các vùng chọn rời nhau khi giữ Ctrl.
        static System.Windows.Forms.Timer poll;
        static bool polling;
        static void StartPoll()
        {
            try
            {
                poll = new System.Windows.Forms.Timer(); poll.Interval = 70;
                poll.Tick += delegate
                {
                    if (polling || App == null) return;
                    polling = true;
                    try { new WordTables(App).Poll(); } catch { }
                    finally { polling = false; }
                };
                poll.Start();
                Log.W("Host: da bat bo theo doi vung chon (Ctrl)");
            }
            catch (Exception ex) { Log.W("StartPoll loi: " + ex); }
        }
        public static void StopPoll() { try { if (poll != null) { poll.Stop(); poll.Dispose(); poll = null; } } catch { } }

        public static T OnWord<T>(Func<T> f)
        {
            if (sync == null || !sync.IsHandleCreated || Thread.CurrentThread.ManagedThreadId == mainId) return f();
            T result = default(T); Exception err = null;
            using (var done = new ManualResetEvent(false))
            {
                sync.BeginInvoke(new MethodInvoker(delegate
                {
                    try { result = f(); }
                    catch (Exception e) { err = e; }
                    finally { try { done.Set(); } catch { } }
                }));
                if (!done.WaitOne(180000)) throw new UserMsg("Word không phản hồi sau 3 phút. Hãy kiểm tra xem Word có đang hiện hộp thoại chờ trả lời không.");
            }
            if (err != null) throw err;
            return result;
        }

        public static string Explain(Exception ex)
        {
            if (ex is UserMsg) return ex.Message;
            var ce = ex as COMException;
            if (ce != null)
            {
                uint code = unchecked((uint)ce.ErrorCode);
                if (code == 0x80010001 || code == 0x8001010A) return "Word đang bận (có hộp thoại đang mở). Hãy đóng hộp thoại trong Word rồi thử lại.";
                return "Word báo lỗi: " + ce.Message;
            }
            if (ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) return "Lỗi gọi Word (có thể do phiên bản Word): " + ex.Message;
            return "Lỗi: " + ex.Message;
        }
    }

    [ComVisible(true)]
    [Guid("6F1C2B54-9A3E-4D0B-8B7A-3C5D2E1F0A47")]
    [ProgId("TroLyKyThuat.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        static Connect() { Log.W("Connect: static constructor (DLL da nap)"); }
        public Connect() { Log.W("Connect: constructor"); }

        object wordApp;
        PaneForm pane;
        Thread uiThread;
        readonly ManualResetEvent ready = new ManualResetEvent(false);

        const string RibbonXmlRaw =
@"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabTools"" label=""Công cụ"">
        <group id=""grpTlFont"" label=""Chữ"">
          <control idMso=""FormatPainter"" size=""large"" />
          <box id=""bxTlFont"" boxStyle=""vertical"">
            <box id=""bxTlFont1"" boxStyle=""horizontal"">
              <control idMso=""Font"" />
              <control idMso=""FontSize"" />
            </box>
            <box id=""bxTlFont2"" boxStyle=""horizontal"">
              <control idMso=""Bold"" showLabel=""false"" />
              <control idMso=""Italic"" showLabel=""false"" />
              <control idMso=""Underline"" showLabel=""false"" />
              <control idMso=""Strikethrough"" showLabel=""false"" />
              <control idMso=""Subscript"" showLabel=""false"" />
              <control idMso=""Superscript"" showLabel=""false"" />
            </box>
            <box id=""bxTlFont3"" boxStyle=""horizontal"">
              <control idMso=""FontColorPicker"" showLabel=""false"" />
              <control idMso=""TextHighlightColorPicker"" showLabel=""false"" />
              <control idMso=""ChangeCase"" showLabel=""false"" />
              <control idMso=""ClearFormatting"" showLabel=""false"" />
            </box>
          </box>
        </group>
        <group id=""grpTlPara"" label=""Đoạn văn"">
          <button id=""tbParaFormat"" label=""Chuẩn hóa đoạn văn"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Chuẩn hóa đoạn tại con trỏ hoặc các đoạn bôi đen"" supertip=""Giống nút Chuẩn hóa đoạn văn ở thẻ Chuẩn hóa: canh lề, cỡ chữ, giãn dòng, khoảng cách và thụt dòng theo Cài đặt. Bôi đen nhiều đoạn thì áp cho tất cả cùng lúc."" />
          <box id=""bxTlPara"" boxStyle=""vertical"">
            <box id=""bxTlPara1"" boxStyle=""horizontal"">
              <control idMso=""Bullets"" showLabel=""false"" />
              <control idMso=""Numbering"" showLabel=""false"" />
              <control idMso=""MultilevelListGallery"" showLabel=""false"" />
              <control idMso=""IndentDecrease"" showLabel=""false"" />
              <control idMso=""IndentIncrease"" showLabel=""false"" />
              <control idMso=""ParagraphMarks"" showLabel=""false"" />
            </box>
            <box id=""bxTlPara2"" boxStyle=""horizontal"">
              <control idMso=""AlignLeft"" showLabel=""false"" />
              <control idMso=""AlignCenter"" showLabel=""false"" />
              <control idMso=""AlignRight"" showLabel=""false"" />
              <control idMso=""AlignJustify"" showLabel=""false"" />
              <control idMso=""LineSpacingMenu"" showLabel=""false"" />
            </box>
            <box id=""bxTlPara3"" boxStyle=""horizontal"">
              <control idMso=""ParagraphDialog"" />
            </box>
          </box>
        </group>
        <group id=""grpTlPage"" label=""Trang"">
          <button id=""tbPageSetup"" label=""Chuẩn hóa dàn trang"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Chuẩn hóa khổ giấy và lề trang"" supertip=""Giống nút Chuẩn hóa dàn trang ở thẻ Chuẩn hóa: khổ giấy, lề trang dọc và ngang, đầu và chân trang theo Cài đặt."" />
          <box id=""bxTlPage1"" boxStyle=""vertical"">
            <button id=""tbAddPort"" label=""Thêm trang dọc"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Thêm 1 trang dọc ngay dưới trang hiện tại"" />
            <button id=""tbAddLand"" label=""Thêm trang ngang"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Thêm 1 trang ngang ngay dưới trang hiện tại"" />
            <control idMso=""PageSetupDialog"" />
          </box>
          <box id=""bxTlPage2"" boxStyle=""vertical"">
            <control idMso=""PageSizeGallery"" />
            <control idMso=""PageOrientationGallery"" />
            <control idMso=""ColumnsDialog"" />
          </box>
          <box id=""bxTlPage3"" boxStyle=""vertical"">
            <control idMso=""BreaksGallery"" />
            <control idMso=""BordersShadingDialog"" />
            <control idMso=""PageNumberFormat"" />
          </box>
        </group>
        <group id=""grpTlIns"" label=""Chèn"">
          <box id=""bxTlIns1"" boxStyle=""vertical"">
            <control idMso=""SymbolInsertGallery"" />
            <control idMso=""ShapesInsertGallery"" />
            <control idMso=""SmartArtInsert"" />
          </box>
          <box id=""bxTlIns2"" boxStyle=""vertical"">
            <control idMso=""CaptionInsert"" />
            <control idMso=""CrossReferenceInsert"" />
            <control idMso=""TextBoxInsertGallery"" />
          </box>
          <control idMso=""EquationInsertGallery"" size=""large"" />
        </group>
        <group id=""grpTlPic"" label=""Ảnh và xem"">
          <box id=""bxTlPic1"" boxStyle=""vertical"">
            <control idMso=""PicturesCompress"" />
            <control idMso=""PictureCrop"" />
          </box>
          <box id=""bxTlView"" boxStyle=""vertical"">
            <control idMso=""ZoomOnePage"" />
            <control idMso=""ZoomTwoPages"" />
            <control idMso=""ZoomPageWidth"" />
          </box>
        </group>
        <group id=""grpTlTbl"" label=""Bảng"">
          <control idMso=""TableInsertGallery"" size=""large"" />
          <box id=""bxTlTbl1"" boxStyle=""vertical"">
            <control idMso=""ShadingColorPicker"" />
            <control idMso=""TableBordersMenu"" />
            <control idMso=""TextDirectionGallery"" />
          </box>
          <box id=""bxTlTbl2"" boxStyle=""vertical"">
            <control idMso=""TableRowsDistribute"" />
            <control idMso=""TableColumnsDistribute"" />
            <control idMso=""TableRepeatHeaderRows"" />
          </box>
          <box id=""bxTlTbl3"" boxStyle=""vertical"">
            <control idMso=""SplitCells"" />
            <control idMso=""TableSplitTable"" />
            <control idMso=""MergeCells"" />
          </box>
          <box id=""bxTlTbl4"" boxStyle=""vertical"">
            <control idMso=""ConvertTextToTable"" />
            <control idMso=""TableAutoFitMenu"" />
            <control idMso=""TableFormulaDialog"" />
          </box>
          <box id=""bxTlAl"" boxStyle=""vertical"">
            <box id=""bxTlAl1"" boxStyle=""horizontal"">
              <control idMso=""TableCellAlignTopLeft"" showLabel=""false"" />
              <control idMso=""TableCellAlignTopCenter"" showLabel=""false"" />
              <control idMso=""TableCellAlignTopRight"" showLabel=""false"" />
            </box>
            <box id=""bxTlAl2"" boxStyle=""horizontal"">
              <control idMso=""TableCellAlignMiddleLeft"" showLabel=""false"" />
              <control idMso=""TableCellAlignMiddleCenter"" showLabel=""false"" />
              <control idMso=""TableCellAlignMiddleRight"" showLabel=""false"" />
            </box>
            <box id=""bxTlAl3"" boxStyle=""horizontal"">
              <control idMso=""TableCellAlignBottomLeft"" showLabel=""false"" />
              <control idMso=""TableCellAlignBottomCenter"" showLabel=""false"" />
              <control idMso=""TableCellAlignBottomRight"" showLabel=""false"" />
            </box>
          </box>
        </group>
      </tab>
      <tab id=""tabTable"" getLabel=""GetTabTableLabel"">
        <group id=""grpStyle"" label=""Bảng"">
          <button id=""btnApplyAll"" label=""Mọi bảng"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Áp dụng master style cho mọi bảng"" supertip=""Rà soát toàn bộ thông số đang thiết lập (tiêu đề, viền, tô xen kẽ, font, cỡ chữ, cách đoạn, căn ô) và áp dụng cho mọi bảng trong tài liệu. Chỉ ghi thuộc tính đang khác nên chạy nhanh."" />
          <button id=""btnApplySel"" label=""Bảng này"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Áp dụng master style cho bảng đang chọn"" supertip=""Rà soát toàn bộ thông số đang thiết lập và áp dụng cho bảng đang chọn. Các bảng khác giữ nguyên."" />
          <splitButton id=""spLayout"" size=""large"">
            <button id=""btnLayout"" label=""Lặp tiêu đề"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Bật hoặc tắt lặp dòng tiêu đề ở đầu mỗi trang"" supertip=""Áp dụng cho mọi bảng nằm trong phần đang chọn: đặt con trỏ trong một bảng thì chỉ bảng đó, bôi đen từ bảng này sang bảng khác thì tất cả các bảng ở giữa. Bấm lần nữa để tắt. Mũi tên bên dưới: bật hoặc tắt cho mọi bảng."" />
            <menu id=""mnuLayout"" label=""Lặp tiêu đề"" screentip=""Phạm vi lặp tiêu đề"">
              <button id=""btnLayoutAll"" label=""Bật cho mọi bảng"" onAction=""OnBtn"" />
              <button id=""btnLayoutOff"" label=""Tắt cho mọi bảng"" onAction=""OnBtn"" />
            </menu>
          </splitButton>
          <box id=""bxTbl"" boxStyle=""vertical"">
            <box id=""bxHdr"" boxStyle=""horizontal"">
              <labelControl id=""lblHdr"" label=""Tiêu đề"" />
              <button id=""btnHdrBg"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Màu nền dòng tiêu đề"" supertip=""Mở bảng màu thống nhất. Áp dụng ngay cho bảng đang chọn."" />
              <button id=""btnHdrFg"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Màu chữ dòng tiêu đề"" supertip=""Mở bảng màu thống nhất. Áp dụng ngay cho bảng đang chọn."" />
              <menu id=""mnuHdrAl"" showLabel=""false"" getImage=""GetImg"" screentip=""Căn chữ dòng tiêu đề"">
                <button id=""hdrAlL"" label=""Căn trái"" onAction=""OnBtn"" />
                <button id=""hdrAlC"" label=""Căn giữa"" onAction=""OnBtn"" />
                <button id=""hdrAlR"" label=""Căn phải"" onAction=""OnBtn"" />
              </menu>
            </box>
            <box id=""bxBd"" boxStyle=""horizontal"">
              <labelControl id=""lblBd"" label=""Viền"" />
              <comboBox id=""cbBd"" sizeString=""0.25"" maxLength=""5"" getText=""GetCb"" onChange=""OnCb"" getItemCount=""BdCount"" getItemLabel=""BdLabel"" screentip=""Độ dày viền (pt)"" supertip=""Chọn nhanh (0.25; 0.5; 0.75; 1.0; 1.5; 2.0; 3.0) hoặc nhập tay từ 0.25 đến 6.0 với bước nhảy 0.25 pt. Áp dụng ngay cho bảng đang chọn."" />
              <button id=""btnBdClr"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Màu viền"" supertip=""Mở bảng màu thống nhất. Áp dụng ngay cho bảng đang chọn."" />
              <menu id=""mnuBdSt"" showLabel=""false"" getImage=""GetImg"" screentip=""Kiểu nét viền"">
                <button id=""bdStSolid"" label=""Liền"" onAction=""OnBtn"" />
                <button id=""bdStDash"" label=""Đứt"" onAction=""OnBtn"" />
                <button id=""bdStDot"" label=""Chấm"" onAction=""OnBtn"" />
              </menu>
            </box>
            <box id=""bxZb"" boxStyle=""horizontal"">
              <checkBox id=""chkZebra"" label=""Tô xen kẽ"" getPressed=""GetPressed"" onAction=""OnCheck"" screentip=""Tô màu xen kẽ các dòng chẵn"" supertip=""Bật hoặc tắt tô dòng chẵn. Áp dụng ngay cho bảng đang chọn."" />
              <button id=""btnZebraClr"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Màu dòng chẵn"" supertip=""Mở bảng màu thống nhất. Áp dụng ngay cho bảng đang chọn."" />
            </box>
          </box>
          <box id=""bxTblAl"" boxStyle=""vertical"">
            <button id=""btnTblAlL"" label=""Bảng trái"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Canh toàn bảng sang trái lề trang"" supertip=""Đưa toàn bộ bảng đang chọn về sát lề trái trang giấy. Áp dụng ngay."" />
            <button id=""btnTblAlC"" label=""Bảng giữa"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Canh toàn bảng vào giữa lề trang"" supertip=""Căn giữa toàn bộ bảng đang chọn so với lề trang giấy. Áp dụng ngay."" />
            <button id=""btnTblAlR"" label=""Bảng phải"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Canh toàn bảng sang phải lề trang"" supertip=""Đưa toàn bộ bảng đang chọn về sát lề phải trang giấy. Áp dụng ngay."" />
          </box>
        </group>
        <group id=""grpChars"" label=""Ký tự trong bảng"">
          <box id=""bxChar1"" boxStyle=""vertical"">
            <gallery id=""galFont"" size=""normal"" getLabel=""GetFontLabel"" getImage=""GetImg"" columns=""1"" itemWidth=""170"" itemHeight=""26"" showItemLabel=""false"" getItemCount=""FontCount"" getItemImage=""FontImg"" getItemLabel=""FontLabel"" getItemScreentip=""FontLabel"" onAction=""OnFontGal"" screentip=""Font chữ trong bảng"" supertip=""Chọn Times New Roman, Arial, Calibri hoặc Segoe UI (hiển thị đúng kiểu dáng font). Áp dụng ngay cho bảng đang chọn."" />
            <box id=""bxSz"" boxStyle=""horizontal"">
              <labelControl id=""lblSz"" label=""Cỡ chữ"" />
              <comboBox id=""cbSize"" sizeString=""99.9"" maxLength=""5"" getText=""GetCb"" onChange=""OnCb"" getItemCount=""SizeCount"" getItemLabel=""SizeLabel"" screentip=""Cỡ chữ trong bảng (pt)"" supertip=""Chọn cỡ chữ chuẩn Word (8 đến 72) hoặc nhập tay từ 6 đến 72 pt. Áp dụng ngay cho bảng đang chọn."" />
            </box>
            <box id=""bxSp"" boxStyle=""horizontal"">
              <labelControl id=""lblSp"" label=""Cách đoạn"" />
              <editBox id=""edSpace"" sizeString=""99,9"" maxLength=""4"" getText=""GetCb"" onChange=""OnCb"" screentip=""Khoảng cách trước và sau đoạn trong ô (pt)"" supertip=""Gõ số từ 0 đến 12 hoặc bấm mũi tên lên, xuống (mỗi lần 1 pt). Áp dụng ngay cho bảng đang chọn."" />
              <button id=""btnSpInc"" label=""Tăng khoảng cách đoạn"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tăng 1 pt"" />
              <button id=""btnSpDec"" label=""Giảm khoảng cách đoạn"" showLabel=""false"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Giảm 1 pt"" />
            </box>
          </box>
          <box id=""bxChar2"" boxStyle=""vertical"">
            <menu id=""mnuTxtAl"" label=""Căn ô chữ"" size=""normal"" getImage=""GetImg"" screentip=""Căn các ô chứa chữ"">
              <button id=""txtAlL"" label=""Trái"" onAction=""OnBtn"" />
              <button id=""txtAlJ"" label=""Hai bên"" onAction=""OnBtn"" />
            </menu>
            <menu id=""mnuNumAl"" label=""Căn ô số"" size=""normal"" getImage=""GetImg"" screentip=""Căn các ô chứa số"">
              <button id=""numAlL"" label=""Trái"" onAction=""OnBtn"" />
              <button id=""numAlC"" label=""Giữa"" onAction=""OnBtn"" />
              <button id=""numAlR"" label=""Phải"" onAction=""OnBtn"" />
            </menu>
          </box>
        </group>
        <group id=""grpRowCol"" label=""Hàng và cột"">
          <comboBox id=""cbN"" label=""Số lượng"" sizeString=""99"" maxLength=""2"" getText=""GetN"" onChange=""OnN"" getItemCount=""NCount"" getItemLabel=""NLabel"" screentip=""Số hàng hoặc cột thêm cạnh mỗi hàng hoặc cột đang chọn"" supertip=""Chọn 0 đến 3 hoặc nhập tay số từ 0 đến 20. Giữ Ctrl để chọn nhiều hàng, cột rời nhau trong Word."" />
          <menu id=""mnuAddRow"" label=""Thêm hàng"" size=""large"" getImage=""GetImg"" screentip=""Thêm hàng cạnh mỗi hàng đã chọn"">
            <button id=""btnRowAbove"" label=""Phía trên"" getImage=""GetImg"" onAction=""OnBtn"" />
            <button id=""btnRowBelow"" label=""Phía dưới"" getImage=""GetImg"" onAction=""OnBtn"" />
          </menu>
          <menu id=""mnuAddCol"" label=""Thêm cột"" size=""large"" getImage=""GetImg"" screentip=""Thêm cột cạnh mỗi cột đã chọn"">
            <button id=""btnColLeft"" label=""Bên trái"" getImage=""GetImg"" onAction=""OnBtn"" />
            <button id=""btnColRight"" label=""Bên phải"" getImage=""GetImg"" onAction=""OnBtn"" />
          </menu>
          <menu id=""mnuDel"" label=""Xóa"" size=""large"" getImage=""GetImg"" screentip=""Xóa các hàng hoặc cột đã chọn"">
            <button id=""btnDelRow"" label=""Hàng"" getImage=""GetImg"" onAction=""OnBtn"" />
            <button id=""btnDelCol"" label=""Cột"" getImage=""GetImg"" onAction=""OnBtn"" />
          </menu>
        </group>
        <group id=""grpCalc"" label=""Tính trong bảng"">
          <box id=""bxCalc1"" boxStyle=""vertical"">
            <button id=""btnSum"" label=""Tổng"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tính các ô đang chọn"" supertip=""Chọn dãy ô nằm ngang thì kết quả điền vào cột bên phải, dãy dọc thì điền vào hàng bên dưới. Tự nhận hướng, không cần chỉnh."" />
            <button id=""btnAvg"" label=""Trung bình"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tính các ô đang chọn"" supertip=""Chọn dãy ô nằm ngang thì kết quả điền vào cột bên phải, dãy dọc thì điền vào hàng bên dưới. Tự nhận hướng, không cần chỉnh."" />
            <button id=""btnMin"" label=""Nhỏ nhất"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tính các ô đang chọn"" supertip=""Chọn dãy ô nằm ngang thì kết quả điền vào cột bên phải, dãy dọc thì điền vào hàng bên dưới. Tự nhận hướng, không cần chỉnh."" />
          </box>
          <box id=""bxCalc2"" boxStyle=""vertical"">
            <button id=""btnMax"" label=""Lớn nhất"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tính các ô đang chọn"" supertip=""Chọn dãy ô nằm ngang thì kết quả điền vào cột bên phải, dãy dọc thì điền vào hàng bên dưới. Tự nhận hướng, không cần chỉnh."" />
            <button id=""btnCnt"" label=""Đếm số"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tính các ô đang chọn"" supertip=""Chọn dãy ô nằm ngang thì kết quả điền vào cột bên phải, dãy dọc thì điền vào hàng bên dưới. Tự nhận hướng, không cần chỉnh."" />
          </box>
          <box id=""bxCalc3"" boxStyle=""vertical"">
            <button id=""btnFmtCalc"" label=""Tổng theo định dạng"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tổng chỉ các ô có định dạng đã chọn (Rev0.9.1A thử nghiệm)"" supertip=""Chỉ cộng các ô khớp kiểu lọc đặt trong Cài đặt (đậm, tô sáng, màu chữ hoặc màu nền). Các ô số khác bị bỏ qua."" />
            <button id=""btnFmtProd"" label=""Tích theo định dạng"" size=""normal"" showLabel=""true"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tích chỉ các ô có định dạng đã chọn trong từng hàng"" supertip=""Chỉ nhân các ô khớp kiểu lọc đặt trong Cài đặt (đậm, tô sáng, màu chữ hoặc màu nền) trong cùng hàng. Đặt con trỏ trong một ô, hoặc bôi đen một cột nhiều ô để tính tích cho từng hàng. Kết quả ghi vào ô được chọn."" />
          </box>
        </group>
      </tab>
      <tab id=""tabNorm"" label=""Chuẩn hóa"">
        <group id=""grpLayoutDoc"" label=""Dàn trang"">
          <button id=""btnPageSetup"" label=""Chuẩn hóa dàn trang"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Chuẩn hóa khổ giấy và lề trang"" supertip=""Đặt khổ giấy (mặc định A4), lề trang dọc (2-2-3-2 cm), lề trang ngang (3-2-2-2 cm), đầu/chân trang (1 cm). Tùy chỉnh trong Cài đặt."" />
          <button id=""btnParaFormat"" label=""Chuẩn hóa đoạn văn"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Chuẩn hóa đoạn tại con trỏ, căn đều hai bên"" supertip=""Áp dụng riêng cho đoạn chứa con trỏ: căn đều hai bên, Times New Roman -> 13 pt, Arial -> 11 pt, font khác giữ nguyên cỡ; giãn dòng, khoảng cách và thụt dòng theo Cài đặt."" />
          <button id=""btnAddLand"" label=""Thêm trang ngang"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Thêm 1 trang ngang ngay dưới trang hiện tại"" supertip=""Chèn 1 trang khổ ngang ngay dưới trang đặt con trỏ; các trang ngay bên dưới vẫn giữ nguyên quy cách ban đầu."" />
          <button id=""btnAddPort"" label=""Thêm trang dọc"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Thêm 1 trang dọc ngay dưới trang hiện tại"" supertip=""Chèn 1 trang khổ dọc ngay dưới trang đặt con trỏ; các trang ngay bên dưới vẫn giữ nguyên quy cách ban đầu."" />
        </group>
        <group id=""grpNum"" label=""Định dạng số"">
          <button id=""btnNumStyle"" getLabel=""GetLabel"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" getScreentip=""GetTip"" getSupertip=""GetSuper"" />
          <button id=""btnDecInc"" label=""Thêm số lẻ"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Tăng một chữ số thập phân"" supertip=""Áp dụng cho mọi số đang bôi đen: ô trong bảng hoặc số trong văn bản. Tối đa 5 chữ số thập phân. Vùng bôi đen được giữ nguyên để bấm tiếp."" />
          <button id=""btnDecDec"" label=""Bớt số lẻ"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Giảm một chữ số thập phân"" />
          <splitButton id=""spSync"" size=""large"">
            <button id=""btnSync"" getLabel=""GetSyncLabel"" getImage=""GetImg"" onAction=""OnBtn"" getScreentip=""GetSyncTip"" supertip=""Mọi số đang bôi đen có đúng chừng này chữ số thập phân. Số có ít hơn được thêm số 0, số có nhiều hơn được làm tròn. Số nhỏ ở góc biểu tượng là số chữ số thập phân sẽ áp dụng. Bấm mũi tên bên dưới để chọn từ 0 đến 5 chữ số, hoặc chọn kiểu kỹ thuật (2,3 × 10⁶) rồi chỉnh số chữ số ở Cài đặt."" />
            <menu id=""mnuSync"" label=""Đồng bộ số lẻ"" screentip=""Chọn số chữ số thập phân"">
              <button id=""sync0"" label=""0 chữ số thập phân (số nguyên)"" onAction=""OnBtn"" />
              <button id=""sync1"" label=""1 chữ số thập phân"" onAction=""OnBtn"" />
              <button id=""sync2"" label=""2 chữ số thập phân"" onAction=""OnBtn"" />
              <button id=""sync3"" label=""3 chữ số thập phân"" onAction=""OnBtn"" />
              <button id=""sync4"" label=""4 chữ số thập phân"" onAction=""OnBtn"" />
              <button id=""sync5"" label=""5 chữ số thập phân"" onAction=""OnBtn"" />
              <menuSeparator id=""sepSync"" />
              <button id=""syncSci"" label=""Viết kiểu kỹ thuật (2,3 × 10⁶)"" onAction=""OnBtn"" />
            </menu>
          </splitButton>
          <button id=""btnRead"" label=""Đọc số"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Viết số và đơn vị bằng chữ ngay sau số"" supertip=""Đặt con trỏ trong hoặc sát một số rồi bấm: chữ được viết ngay sau số, trong ngoặc, ví dụ 1.234 (một nghìn hai trăm ba mươi bốn). Đơn vị đứng ngay sau số cũng được đọc, ví dụ 20 kPa (hai mươi kilôpascan), 1.000.000 VNĐ (một triệu Việt Nam đồng); tên đơn vị theo Nghị định 134/2007/NĐ-CP, đổi sang tên thông dụng trong Cài đặt. Bôi đen nhiều số thì đọc từng số. Số thập phân đọc từng chữ số sau chữ phẩy, số âm đọc âm, số kỹ thuật 2,3 × 10⁶ đọc hai phẩy ba nhân mười mũ sáu."" />
        </group>
        <group id=""grpRv"" label=""Chuẩn hóa kỹ thuật"">
          <button id=""btnQuick"" label=""Chuẩn hóa"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Chuẩn hóa chữ tại con trỏ hoặc phần bôi đen"" supertip=""Đặt con trỏ vào hoặc bôi đen một từ, công thức bị lỗi mã (ví dụ R_cyl, R_{cyl}, t_req, \Delta P, \frac{a}{b}) rồi bấm. Sửa ngay về dạng chuẩn: chữ in nghiêng, chỉ số dưới đứng thẳng, bỏ các dấu _ { } thừa. Không cần quét cả văn bản."" />
          <button id=""btnReview"" label=""Rà soát"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Rà soát toàn văn bản"" supertip=""Quét công thức hóa học, toán, ký hiệu, đơn vị và số kỹ thuật, rồi duyệt từng mục để sửa."" />
          <button id=""btnIndex"" label=""Kiểu chỉ số"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Đổi kiểu chỉ số của công thức đang bôi đen"" supertip=""Bôi đen công thức hoặc cả đoạn rồi bấm: đổi luân phiên giữa Equation và font thường có chỉ số trên, dưới. Bấm lần nữa để quay lại."" />
          <button id=""btnCompose"" label=""Phương trình"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Soạn phương trình"" supertip=""Gõ phương trình hóa học hoặc toán, tự nhận loại, cảnh báo lỗi, xem trước rồi chèn tại con trỏ."" />
        </group>
        <group id=""grpLang"" label=""Công cụ AI"">
          <button id=""btnTerm"" label=""Thuật ngữ"" size=""normal"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Rà soát và tra thuật ngữ bằng AI"" />
          <button id=""btnStyle"" label=""Văn phong"" size=""normal"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Rà soát văn phong bằng AI"" />
          <button id=""btnGram"" label=""Ngữ pháp"" size=""normal"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Rà soát ngữ pháp bằng AI"" />
          <button id=""btnSyn"" label=""Từ đồng nghĩa"" size=""normal"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Gợi ý từ đồng nghĩa bằng AI"" />
          <button id=""btnAiTools"" label=""Viết lại"" size=""normal"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Viết lại và dịch vùng bôi đen bằng AI"" />
        </group>
        <group id=""grpPane"" label=""Cài đặt"">
          <button id=""btnPane"" label=""Bảng điều khiển"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Mở bảng điều khiển"" supertip=""Rà soát và soạn phương trình."" />
          <button id=""btnSettings"" label=""Cài đặt"" size=""large"" getImage=""GetImg"" onAction=""OnBtn"" screentip=""Cài đặt và trợ giúp"" supertip=""Quy ước số, xóa định dạng chỉnh tay khi áp dụng master style, khóa API, hướng dẫn sử dụng."" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        static string RibbonXml
        {
            get
            {
                string x = RibbonXmlRaw;
                foreach (string[] c in ClrIds) x = x.Replace("<!--P:" + c[1] + "-->", TableParams.PresetMenuXml(c[1], c[0]));
                return x;
            }
        }

        public static string BuildRibbonXml() { return RibbonXml; }
        internal static string RibbonXmlForTest() { return RibbonXml; }

        // ---------------------------------------------------------------- IRibbonExtensibility
        string IRibbonExtensibility.GetCustomUI(string RibbonID) { Log.W("GetCustomUI " + RibbonID); return RibbonXml; }

        // Ribbon gọi các hàm này qua IDispatch (trên luồng chính của Word).
        public string GetN(object control)
        {
            try { return AppSettings.Load().AddN.ToString(CultureInfo.InvariantCulture); } catch { return "1"; }
        }

        object ribbonUi;
        public void OnLoad(object ribbon) { ribbonUi = ribbon; try { WordPalette.LoadFrom(wordApp); } catch { } StartTableDetect(); }
        void Refresh(string id) { try { if (ribbonUi != null) ((dynamic)ribbonUi).InvalidateControl(id); } catch (Exception ex) { Log.W("Invalidate loi: " + ex.Message); } }

        // ---- phát hiện con trỏ trong bảng để gợi ý thẻ "My table ●"
        System.Windows.Forms.Timer tblTimer;
        bool lastInTable;
        bool tblDetecting;
        void StartTableDetect()
        {
            try
            {
                if (tblTimer != null) return;
                tblTimer = new System.Windows.Forms.Timer(); tblTimer.Interval = 400;
                tblTimer.Tick += delegate
                {
                    if (tblDetecting) return;
                    tblDetecting = true;
                    try
                    {
                        bool cur = false;
                        try
                        {
                            if (wordApp != null && (int)((dynamic)wordApp).Documents.Count > 0)
                                cur = (bool)((dynamic)wordApp).Selection.Information(12);
                        }
                        catch { }
                        if (cur != lastInTable)
                        {
                            lastInTable = cur;
                            try { if (ribbonUi != null) ((dynamic)ribbonUi).InvalidateControl("tabTable"); } catch { }
                        }
                    }
                    catch { }
                    finally { tblDetecting = false; }
                };
                tblTimer.Start();
            }
            catch { }
        }
        void StopTableDetect() { try { if (tblTimer != null) { tblTimer.Stop(); tblTimer.Dispose(); tblTimer = null; } } catch { } }

        /// <summary>Nhãn thẻ "My table": thêm dấu chấm tròn khi con trỏ nằm trong bảng.</summary>
        public string GetTabTableLabel(object control)
        {
            try { return lastInTable ? "My table \u25CF" : "My table"; }
            catch { return "My table"; }
        }

        /// <summary>Ribbon hỏi icon của từng nút: vẽ bằng mã (không dùng tên imageMso có thể sai).</summary>
        public object GetImg(object control)
        {
            try
            {
                string id = (string)((dynamic)control).Id;
                return Icons.Picture(ImgKey(id));
            }
            catch (Exception ex) { Log.W("GetImg loi: " + ex.Message); return null; }
        }

        /// <summary>Mã icon theo mã điều khiển: nút đổi theo cài đặt (kiểu dấu, số chữ số thập phân, màu) được vẽ lại mỗi khi Ribbon hỏi.</summary>
        static string ImgKey(string id)
        {
            if (id.StartsWith("auto_")) return "sw:" + WordPalette.AutoHex(id.Substring(5));
            switch (id)
            {
                case "btnNumStyle": return id + ":" + AppSettings.Load().NumStyle;
                case "btnSync": { AppSettings q = AppSettings.Load(); return id + ":" + (q.SyncSci ? "sci" : q.SyncDec.ToString(CultureInfo.InvariantCulture)); }
                case "btnHdrBg": case "mnuHdrBg": return "btnHdrBg:" + AppSettings.Load().HdrBg;
                case "btnHdrFg": case "mnuHdrFg": return "btnHdrFg:" + AppSettings.Load().HdrFg;
                case "btnBdClr": case "mnuBdClr": return "btnBdClr:" + AppSettings.Load().BorderColor;
                case "btnZebraClr": case "mnuZebraClr": return "btnZebraClr:" + AppSettings.Load().ZebraColor;
                case "btnSzInc": case "btnBdInc": return "btnNInc";
                case "btnSzDec": case "btnBdDec": return "btnNDec";
            }
            return id;
        }

        public bool GetPressed(object control)
        {
            try { return AppSettings.Load().Zebra; } catch { return false; }
        }

        public void OnCheck(object control, bool pressed)
        {
            try
            {
                string id = (string)((dynamic)control).Id;
                if (id != "chkZebra") return;
                AppSettings s = AppSettings.Load(); s.Zebra = pressed; s.Save();
                Notify((pressed ? "Đã bật tô xen kẽ dòng chẵn." : "Đã tắt tô xen kẽ.") + TableParams.Hint);
                ApplyNow("zebra");
            }
            catch (Exception ex) { Log.W("OnCheck loi: " + ex); }
        }

        public string GetLabel(object control)
        {
            try { return "Dấu " + NumberEdit.StyleShort(AppSettings.Load().NumStyle); } catch { return "Dấu số"; }
        }
        public string GetTip(object control)
        {
            try
            {
                string cur = AppSettings.Load().NumStyle;
                return "Kiểu số hiện hành: " + NumberEdit.StyleName(cur) + ". Bấm để đổi sang " + NumberEdit.StyleName(NumberEdit.NextStyle(cur));
            }
            catch { return "Đổi kiểu dấu phân cách"; }
        }
        public string GetSuper(object control)
        {
            return "Đổi mọi số đang bôi đen (ô trong bảng hoặc số trong văn bản) sang kiểu kế tiếp, theo vòng Việt, Mỹ, SI. Kiểu mặc định và dấu hàng nghìn chỉnh ở Cài đặt.";
        }


        public string GetSyncLabel(object control)
        {
            try { AppSettings s = AppSettings.Load(); return s.SyncSci ? "Đồng bộ: kỹ thuật" : "Đồng bộ: " + s.SyncDec + " số lẻ"; } catch { return "Đồng bộ số lẻ"; }
        }
        public string GetSyncTip(object control)
        {
            try
            {
                AppSettings s = AppSettings.Load();
                return s.SyncSci ? "Viết mọi số đang bôi đen theo kiểu kỹ thuật, ví dụ 2,3 × 10⁶ (phần đầu có " + s.SciDec + " chữ số thập phân, chỉnh ở Cài đặt)"
                    : "Mọi số đang bôi đen có đúng " + s.SyncDec + " chữ số thập phân";
            }
            catch { return "Đồng bộ số chữ số thập phân"; }
        }

        // ---- danh sách font (4 font cố định, hiển thị đúng kiểu chữ), cỡ chữ chuẩn Word, độ dày viền, số lượng
        public int FontCount(object control) { return TableParams.FontPresets.Length; }
        public string FontLabel(object control, int index) { return index >= 0 && index < TableParams.FontPresets.Length ? TableParams.FontPresets[index] : ""; }
        public object FontImg(object control, int index)
        {
            try
            {
                string fn = FontLabel(control, index);
                return Icons.FontPicture(fn);
            }
            catch (Exception ex) { Log.W("FontImg loi: " + ex.Message); return null; }
        }
        public string GetFontLabel(object control)
        {
            try { return "Font: " + AppSettings.Load().FontName; } catch { return "Font"; }
        }
        public void OnFontGal(object control, string selectedId, int selectedIndex)
        {
            try
            {
                string fn = FontLabel(control, selectedIndex);
                if (string.IsNullOrEmpty(fn)) return;
                AppSettings s = AppSettings.Load(); string m = TableParams.SetFont(s, fn); s.Save();
                Notify(m); Refresh("galFont"); ApplyNow("font");
            }
            catch (Exception ex) { Log.W("OnFontGal loi: " + ex); }
        }
        public int SizeCount(object control) { return TableParams.SizePresets.Length; }
        public string SizeLabel(object control, int index) { return index >= 0 && index < TableParams.SizePresets.Length ? TableParams.SizePresets[index] : ""; }
        public int BdCount(object control) { return TableParams.BorderPresets.Length; }
        public string BdLabel(object control, int index) { return index >= 0 && index < TableParams.BorderPresets.Length ? TableParams.BorderPresets[index] : ""; }
        public int NCount(object control) { return TableParams.NPresets.Length; }
        public string NLabel(object control, int index) { return index >= 0 && index < TableParams.NPresets.Length ? TableParams.NPresets[index] : ""; }

        // ---- bảng màu giống ô chọn màu của Word
        static string GalId(object control) { return (string)((dynamic)control).Id; }
        public int GalCount(object control)
        {
            try { return GalId(control).StartsWith("gT_") ? WordPalette.ThemeCount : WordPalette.StdCount; } catch { return 0; }
        }
        public object GalImg(object control, int index)
        {
            try
            {
                string t; string hex = WordPalette.Resolve(GalId(control), index, out t);
                return Icons.Picture("gs:" + (hex ?? "CCCCCC"));
            }
            catch (Exception ex) { Log.W("GalImg loi: " + ex.Message); return null; }
        }
        public string GalLabel(object control, int index)
        {
            try { return GalId(control).StartsWith("gT_") ? WordPalette.ThemeLabel(index) : WordPalette.StdLabel(index); } catch { return ""; }
        }
        public void OnGal(object control, string selectedId, int selectedIndex)
        {
            try
            {
                string target; string hex = WordPalette.Resolve(GalId(control), selectedIndex, out target);
                if (hex == null || target == null) return;
                SetTarget(target, hex);
            }
            catch (Exception ex) { Log.W("OnGal loi: " + ex); }
        }
        void SetTarget(string target, string hex)
        {
            AppSettings s = AppSettings.Load(); string m = TableParams.SetColor(target, s, hex); s.Save(); Notify(m); Refresh(MenuOf(target)); ApplyNow(target == "border" ? "borderClr" : target);
        }

        public string GetZebraLabel(object control)
        {
            try { return AppSettings.Load().Zebra ? "Tô xen kẽ: đang bật (bấm để tắt)" : "Tô xen kẽ: đang tắt (bấm để bật)"; } catch { return "Tô xen kẽ"; }
        }

        static string Fmt(double v) { return v.ToString("0.##", CultureInfo.CurrentCulture); }

        public string GetCb(object control)
        {
            try
            {
                string id = (string)((dynamic)control).Id; AppSettings s = AppSettings.Load();
                if (id == "cbFont") return s.FontName;
                if (id == "cbSize" || id == "edSize") return Fmt(s.FontSize);
                if (id == "cbSpace" || id == "edSpace") return Fmt(s.CellSpace);
                if (id == "cbBd" || id == "edBd") return s.BorderPt.ToString("0.0#", CultureInfo.InvariantCulture);
            }
            catch (Exception ex) { Log.W("GetCb loi: " + ex.Message); }
            return "";
        }

        public void OnCb(object control, string text)
        {
            try
            {
                string id = (string)((dynamic)control).Id; AppSettings s = AppSettings.Load(); string m;
                if (id == "cbFont") m = TableParams.SetFont(s, text);
                else if (id == "cbSize" || id == "edSize") m = TableParams.SetSize(s, text);
                else if (id == "cbBd" || id == "edBd") m = TableParams.SetBorder(s, text);
                else m = TableParams.SetSpace(s, text);
                if (m.StartsWith("Đã")) s.Save();
                Notify(m);
                Refresh(id);
                if (m.StartsWith("Đã"))
                {
                    if (id == "cbFont" || id == "cbSize" || id == "edSize") ApplyNow("font");
                    else if (id == "cbBd" || id == "edBd") ApplyNow("borderW");
                    else ApplyNow("space");
                }
            }
            catch (Exception ex) { Log.W("OnCb loi: " + ex); }
        }

        static readonly string[][] ClrIds = { new[] { "clrHdrBg", "hdrBg", "Màu nền dòng tiêu đề", "btnHdrBg" }, new[] { "clrHdrFg", "hdrFg", "Màu chữ dòng tiêu đề", "btnHdrFg" },
            new[] { "clrBorder", "border", "Màu viền", "btnBdClr" }, new[] { "clrZebra", "zebra", "Màu dòng chẵn", "btnZebraClr" } };

        void OpenColorPalette(string target)
        {
            try { WordPalette.LoadFrom(wordApp); } catch { }
            AppSettings s = AppSettings.Load();
            string cur = TableParams.GetColor(target, s);
            using (DpiUtil.EnterPerMonitorV2())
            using (var cpf = new ColorPickerForm(target, cur, PresetColors()))
            {
                cpf.ShowDialog();
                Log.W("Bang mau: " + target + " -> " + (cpf.SelectedHex ?? "(huy)") + ", KetQua=" + cpf.DialogResult);
                // Rev0.9.5: tin vào màu đã chọn (SelectedHex), không phụ thuộc DialogResult vì sự kiện mất tiêu điểm có thể đổi nó.
                if (!string.IsNullOrEmpty(cpf.SelectedHex))
                    SetTarget(target, cpf.SelectedHex);
            }
        }

        static string MenuOf(string target)
        {
            foreach (string[] c in ClrIds) if (c[1] == target) return c[3];
            return "";
        }

        void PickColor(string ribbonId)
        {
            foreach (string[] c in ClrIds)
            {
                if (c[0] != ribbonId) continue;
                AppSettings s = AppSettings.Load();
                using (var d = new ColorDialog())
                {
                    d.FullOpen = true; d.AnyColor = true;
                    string hex = TableParams.GetColor(c[1], s);
                    try { d.Color = System.Drawing.Color.FromArgb(Convert.ToInt32(hex.Substring(0, 2), 16), Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16)); } catch { }
                    d.CustomColors = PresetColors();
                    if (d.ShowDialog() != DialogResult.OK) { Notify("Đã hủy chọn " + c[2].ToLowerInvariant() + "."); return; }
                    string m = TableParams.SetColor(c[1], s, string.Format("{0:X2}{1:X2}{2:X2}", d.Color.R, d.Color.G, d.Color.B));
                    s.Save(); Notify(m); Refresh(c[3]); ApplyNow(c[1] == "border" ? "borderClr" : c[1]);
                }
                return;
            }
        }

        static int[] PresetColors()
        {
            string[] h = { "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "0E2841", "E8E8E8", "FFFFFF", "000000", "F2F2F2", "FFF2CC", "C00000", "FFC000", "00B050", "0070C0" };
            int[] r = new int[h.Length];
            for (int i = 0; i < h.Length; i++) r[i] = AppSettings.HexToWord(h[i]);
            return r;
        }

        public void OnN(object control, string text)
        {
            try
            {
                int n;
                if (!int.TryParse((text ?? "").Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out n) || n < 0 || n > 20)
                {
                    Notify("Hãy nhập một số từ 0 đến 20."); Refresh("cbN"); return;
                }
                AppSettings st = AppSettings.Load(); st.AddN = n; st.Save();
                SayInPane("Số hàng hoặc cột thêm cho mỗi hàng hoặc cột: " + n + ".");
            }
            catch (Exception ex) { Log.W("OnN loi: " + ex); }
        }

        public void OnBtn(object control)
        {
            string id = "";
            try { id = (string)((dynamic)control).Id; } catch (Exception e0) { Log.W("OnBtn: khong doc duoc Id: " + e0); }
            Log.W("OnBtn " + id);
            try
            {
                switch (id)
                {
                    case "btnPageSetup": case "tbPageSetup": Do(delegate(AppSettings s) { return new WordLayout(wordApp).StandardizePageSetup(s); }); break;
                    case "btnParaFormat": case "tbParaFormat": Do(delegate(AppSettings s) { return new WordLayout(wordApp).StandardizeParagraphs(s); }); break;
                    case "btnAddLand": case "tbAddLand": Do(delegate(AppSettings s) { return new WordLayout(wordApp).InsertOrientedPage(s, true); }); break;
                    case "btnAddPort": case "tbAddPort": Do(delegate(AppSettings s) { return new WordLayout(wordApp).InsertOrientedPage(s, false); }); break;
                    case "btnPane": ShowPane(null, false); break;
                    case "btnSettings": ShowSettings(); break;
                    case "btnReview": ShowPane("rv", true); break;
                    case "btnCompose": ShowPane("ct", false); break;
                    case "btnTerm": ShowPane("term", false); break;
                    case "btnStyle": ShowPane("style", false); break;
                    case "btnGram": ShowPane("gram", false); break;
                    case "btnAiTools": ShowPane("ai", false); break;
                    case "btnSyn": ShowPane("syn", false); break;
                    case "btnApplyAll": Exec(delegate(WordTables w, AppSettings s) { return w.ApplyAll(s); }, null); break;
                    case "btnApplySel": Exec(delegate(WordTables w, AppSettings s) { return w.ApplyToSelected(s); }, null); break;
                    case "btnTblAlL": Exec(delegate(WordTables w, AppSettings s) { return w.AlignTables(0); }, null); break;
                    case "btnTblAlC": Exec(delegate(WordTables w, AppSettings s) { return w.AlignTables(1); }, null); break;
                    case "btnTblAlR": Exec(delegate(WordTables w, AppSettings s) { return w.AlignTables(2); }, null); break;
                    case "btnHdrBg": OpenColorPalette("hdrBg"); break;
                    case "btnHdrFg": OpenColorPalette("hdrFg"); break;
                    case "btnBdClr": OpenColorPalette("border"); break;
                    case "btnZebraClr": OpenColorPalette("zebra"); break;
                    case "btnRowAbove": Exec(delegate(WordTables w, AppSettings s) { return w.AddRows(s.AddN, true, ""); }, null); break;
                    case "btnRowBelow": Exec(delegate(WordTables w, AppSettings s) { return w.AddRows(s.AddN, false, ""); }, null); break;
                    case "btnColLeft": Exec(delegate(WordTables w, AppSettings s) { return w.AddCols(s.AddN, true, ""); }, null); break;
                    case "btnColRight": Exec(delegate(WordTables w, AppSettings s) { return w.AddCols(s.AddN, false, ""); }, null); break;
                    case "btnDelRow": Exec(delegate(WordTables w, AppSettings s) { return w.DeleteRows(""); }, null); break;
                    case "btnDelCol": Exec(delegate(WordTables w, AppSettings s) { return w.DeleteCols(""); }, null); break;
                    case "btnLayout": Exec(delegate(WordTables w, AppSettings s) { return w.SetRepeatHeader("sel"); }, null); break;
                    case "btnLayoutAll": Exec(delegate(WordTables w, AppSettings s) { return w.SetRepeatHeader("allOn"); }, null); break;
                    case "btnLayoutOff": Exec(delegate(WordTables w, AppSettings s) { return w.SetRepeatHeader("allOff"); }, null); break;
                    case "btnSum": CalcWith("sum"); break;
                    case "btnAvg": CalcWith("avg"); break;
                    case "btnMin": CalcWith("min"); break;
                    case "btnMax": CalcWith("max"); break;
                    case "btnCnt": CalcWith("cnt"); break;
                    case "btnFmtCalc": CalcFmtWith("sum"); break;
                    case "btnFmtProd": CalcFmtProd(); break;
                    case "btnNumStyle": Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Apply(s, "cycle"); }); Refresh("btnNumStyle"); break;
                    case "btnDecInc": Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Apply(s, "inc"); }); break;
                    case "btnDecDec": Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Apply(s, "dec"); }); break;
                    case "btnSync": Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Apply(s, "sync"); }); break;
                    case "sync0": case "sync1": case "sync2": case "sync3": case "sync4": case "sync5": case "syncSci":
                        {
                            AppSettings s0 = AppSettings.Load();
                            if (id == "syncSci") s0.SyncSci = true; else { s0.SyncSci = false; s0.SyncDec = id[4] - '0'; }
                            s0.Save(); Refresh("btnSync");
                            Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Apply(s, "sync"); });
                            break;
                        }
                    case "btnRead": Do(delegate(AppSettings s) { return new WordNumbers(wordApp).Read(s); }); break;
                    case "btnSzInc": case "btnSzDec": case "btnSpInc": case "btnSpDec": case "btnBdInc": case "btnBdDec":
                        {
                            AppSettings s = AppSettings.Load(); string m = TableParams.Step(id, s); s.Save(); Notify(m);
                            Refresh(id.StartsWith("btnSz") ? "cbSize" : (id.StartsWith("btnSp") ? "edSpace" : "cbBd"));
                            ApplyNow(id.StartsWith("btnSz") ? "font" : (id.StartsWith("btnSp") ? "space" : "borderW"));
                            break;
                        }
                    case "btnIndex": Do(delegate(AppSettings s) { return new WordIndexStyle(wordApp).Cycle(null, s.EqStacked); }); break;
                    case "btnQuick": Do(delegate(AppSettings s) { return new WordReview(wordApp).QuickFix(s); }); break;
                    case "btnNInc": case "btnNDec":
                        {
                            AppSettings s = AppSettings.Load();
                            s.AddN = Math.Max(0, Math.Min(20, s.AddN + (id == "btnNInc" ? 1 : -1))); s.Save();
                            Refresh("cbN"); Notify("Số hàng hoặc cột thêm cho mỗi hàng hoặc cột: " + s.AddN + ".");
                            break;
                        }
                    default:
                        if (id.StartsWith("auto_")) { string tg = id.Substring(5); SetTarget(tg, WordPalette.AutoHex(tg)); }
                        else if (id.StartsWith("clr")) PickColor(id);
                        else
                        {
                            AppSettings s = AppSettings.Load(); string m = TableParams.Apply(id, s);
                            if (m != null)
                            {
                                s.Save(); Notify(m);
                                if (id == "zebraToggle") Refresh("chkZebra");
                                if (id.StartsWith("bdW")) Refresh("cbBd");
                                string grp = id.StartsWith("hdrAl") ? "hdrAl"
                                    : id.StartsWith("bdSt") ? "borderSt"
                                    : id.StartsWith("bdW") ? "borderW"
                                    : (id.StartsWith("txtAl") || id.StartsWith("numAl")) ? "align"
                                    : (id == "zebraToggle") ? "zebra" : null;
                                if (grp != null) ApplyNow(grp);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex) { Log.W("OnBtn loi: " + ex); Notify(Host.Explain(ex)); }
        }

        void CalcWith(string op)
        {
            Exec(delegate(WordTables w, AppSettings s) { return w.Calc(s); }, delegate(AppSettings s) { s.CalcOp = op; s.CalcFmtOn = false; });
        }
        /// <summary>Tính như nút Tổng, nhưng chỉ lấy các ô khớp kiểu định dạng lọc đã cài (Rev0.9.1A).</summary>
        void CalcFmtWith(string op)
        {
            Exec(delegate(WordTables w, AppSettings s) { return w.Calc(s); }, delegate(AppSettings s) { s.CalcOp = op; s.CalcFmtOn = true; });
        }
        void CalcFmtProd()
        {
            Exec(delegate(WordTables w, AppSettings s) { return w.CalcProdFmt(s); }, delegate(AppSettings s) { s.CalcOp = "prod"; s.CalcFmtOn = true; });
        }
        /// <summary>Chạy một lệnh làm việc trên vùng bôi đen (Định dạng số, Kiểu chỉ số) và báo kết quả.</summary>
        void Do(Func<AppSettings, string> job)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string msg;
            try { msg = job(AppSettings.Load()); }
            catch (Exception ex) { Log.W("Do loi: " + ex); msg = Host.Explain(ex); }
            sw.Stop();
            Notify(msg + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.InvariantCulture) + " giây)");
        }

        /// <summary>Chạy một lệnh bảng ngay trên luồng chính của Word (đang ở đúng luồng khi Ribbon gọi).</summary>
        void Exec(Func<WordTables, AppSettings, string> job, Action<AppSettings> pre)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            string msg;
            try
            {
                AppSettings st = AppSettings.Load();
                if (pre != null) { pre(st); st.Save(); }
                msg = job(new WordTables(wordApp), st);
            }
            catch (Exception ex) { Log.W("Exec loi: " + ex); msg = Host.Explain(ex); }
            sw.Stop();
            Notify(msg + " (" + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.InvariantCulture) + " giây)");
        }

        void ApplyNow(string group) { Exec(delegate(WordTables w, AppSettings s) { return w.ApplyNow(s, group); }, null); }

        void Notify(string msg)
        {
            try { ((dynamic)wordApp).StatusBar = "Trợ lý Kỹ thuật: " + msg; } catch { }
            if (SayInPane(msg)) return;
            // Rev0.7: không hiện hộp thoại sau mỗi lệnh (mất thời gian); kết quả nằm trên thanh trạng thái của Word. Bật lại ở Cài đặt nếu cần.
            try { if (AppSettings.Load().ShowDialog) MessageBox.Show(msg, "Trợ lý Kỹ thuật"); } catch { }
        }

        bool SayInPane(string msg)
        {
            try
            {
                if (pane != null && !pane.IsDisposed && pane.Visible)
                {
                    pane.BeginInvoke(new MethodInvoker(delegate { pane.Say(msg); }));
                    return true;
                }
            }
            catch { }
            return false;
        }

        // ---------------------------------------------------------------- IDTExtensibility2
        void IDTExtensibility2.OnConnection(IntPtr Application, int ConnectMode, IntPtr AddInInst, IntPtr custom)
        {
            Log.W("OnConnection bat dau, mode " + ConnectMode);
            try
            {
                if (Application != IntPtr.Zero) wordApp = Marshal.GetObjectForIUnknown(Application);
                AppDomain.CurrentDomain.UnhandledException += delegate(object o, UnhandledExceptionEventArgs e) { Log.W("UNHANDLED: " + e.ExceptionObject); };
                AiSessionUsage.Reset();
                Host.Init(wordApp);
                Log.W("OnConnection OK, CLR " + Environment.Version + ", " + (IntPtr.Size == 8 ? "64-bit" : "32-bit"));
            }
            catch (Exception ex) { Log.W("OnConnection loi: " + ex); }
        }
        void IDTExtensibility2.OnDisconnection(int RemoveMode, IntPtr custom) { Log.W("OnDisconnection " + RemoveMode); Shutdown(); }
        void IDTExtensibility2.OnAddInsUpdate(IntPtr custom) { }
        void IDTExtensibility2.OnStartupComplete(IntPtr custom) { Log.W("OnStartupComplete"); }
        void IDTExtensibility2.OnBeginShutdown(IntPtr custom) { Log.W("OnBeginShutdown"); Shutdown(); }

        // ---------------------------------------------------------------- bảng điều khiển chạy trên luồng STA riêng
        // Lý do: form cần vòng lặp thông điệp riêng để phím Tab, mũi tên, trình đọc màn hình hoạt động đúng.
        // Mọi lệnh gọi Word từ bảng điều khiển đi qua Host.OnWord để chạy ở luồng chính của Word.
        void EnsurePane()
        {
            if (pane != null && !pane.IsDisposed) return;
            ready.Reset();
            uiThread = new Thread(delegate()
            {
                try
                {
                    Log.W("Tao bang dieu khien...");
                    IntPtr prevDpi;
                    DpiUtil.TryEnablePerMonitorV2(out prevDpi);
                    Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                    Application.ThreadException += delegate(object o, ThreadExceptionEventArgs e) { Log.W("ThreadException: " + e.Exception); };
                    Application.EnableVisualStyles();
                    pane = new PaneForm(wordApp);
                    IntPtr h = pane.Handle;   // buộc tạo handle trên luồng này, chưa hiện form
                    Log.W("Bang dieu khien: " + DpiUtil.Describe(pane));
                    if (DpiUtil.WindowAwareness(h) != 2)
                    {
                        // Rev0.9.5.2: cửa sổ không ở chế độ Per-Monitor thì Windows kéo giãn ảnh làm chữ mờ. Bật lại ngữ cảnh DPI rồi tạo lại một lần.
                        Log.W("CANH BAO: bang dieu khien khong o che do DPI theo tung man hinh, tao lai");
                        try { pane.Dispose(); } catch { }
                        DpiUtil.TryEnablePerMonitorV2(out prevDpi);
                        pane = new PaneForm(wordApp);
                        h = pane.Handle;
                        Log.W("Sau khi tao lai: " + DpiUtil.Describe(pane));
                    }
                    Log.W("Bang dieu khien san sang");
                    ready.Set();
                    Application.Run();        // vòng lặp thông điệp, thoát bằng Application.ExitThread()
                }
                catch (Exception ex)
                {
                    Log.W("Loi tao bang dieu khien: " + ex);
                    pane = null; ready.Set();
                    try { MessageBox.Show("Không tạo được bảng điều khiển: " + ex.Message + "\r\nChi tiết trong tệp log.txt ở thư mục %TEMP%\\TroLyKyThuat."); } catch { }
                }
            });
            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.IsBackground = true;
            uiThread.Start();
            ready.WaitOne(8000);
        }

        void ShowSettings()
        {
            EnsurePane();
            if (pane == null) return;
            pane.BeginInvoke(new MethodInvoker(delegate { pane.ShowSettingsDialog(); }));
        }

        void ShowPane(string tab, bool scan)
        {
            EnsurePane();
            if (pane == null) return;
            pane.BeginInvoke(new MethodInvoker(delegate { pane.ShowAndFocus(tab, scan); }));
        }

        void Shutdown()
        {
            Host.StopPoll();
            StopTableDetect();
            try
            {
                if (pane != null && !pane.IsDisposed)
                {
                    pane.BeginInvoke(new MethodInvoker(delegate { pane.AllowClose = true; pane.Close(); Application.ExitThread(); }));
                    if (uiThread != null) uiThread.Join(1500);
                }
            }
            catch { }
            pane = null;
            try { if (wordApp != null && Marshal.IsComObject(wordApp)) Marshal.ReleaseComObject(wordApp); } catch { }
            wordApp = null;
        }
    }
}
