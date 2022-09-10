# Cài đặt trước khi chạy scripts
- Microsoft Excel (dùng để chạy `ExcelToPDF.ps1`)
- [pdftk](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/) (dùng để chạy `EncryptPDF.ps1`)
- Sau khi cài `pdftk`, thêm đường dẫn tới folder `\bin` của PDFTk (thường là `C:\Program Files (x86)\PDFtk\bin`) vào `PATH Environment Variables`. Ví dụ như trong ảnh này ![PATH Environment Variables Guideline](https://helpdeskgeek.com/wp-content/pictures/2017/09/edit-environment-variables.png)
- [Guideline add PATH Environment Variables](https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee537574(v=office.14)#to-add-a-path-to-the-path-environment-variable) 
> 1. On the Start menu, right-click Computer.
> 1. In the System dialog box, click Advanced system settings.
> 1. On the Advanced tab of the System Properties dialog box, click Environment Variables.
> 1. In the System Variables box of the Environment Variables dialog box, scroll to Path and select it.

# Bulk Export Excels to PDF
- Copy các file Excels vào cùng folder với file `./ExcelToPDF.ps`.
- Chọn file `./ExcelToPDF.ps1`, nhấn chuột phải và chọn `Run with Powershell`.
- Enter "Y" nếu được hỏi "Execution Policy Changes".
- Sau khi chạy xong script, các file PDF sẽ được tạo từ các file Excels cùng tên, trong cùng folder.

# Bulk set password to PDF
- Mở file Passwords.csv bằng Excel. Điền tên các file Excel và password.của từng file vào trong 2 cột Name và Password và lưu lại thay đổi. 
- Chọn file `./EncyptPDF.ps1` nhấn chuột phải và chọn `Run with Powershell`.
- Enter "Y" nếu được hỏi "Execution Policy Changes".
- Khi được hỏi `"Enter an owner password"`, nhập owner password. 
> Owner password được dùng để remove user password khỏi các file PDF.
> 
> Nếu không set Owner password thì ngay cả khi file PDF được bảo vệ bởi user password, người dùng vẫn có thể remove user password khỏi file PDF được.
- Sau khi chạy xong script, các file PDF có set password sẽ được tạo trong folder `\output` ở cùng thư mục.
