function Get-MainFormTexts {
    return [ordered]@{
        WindowTitle          = 'ADUserCreator Desktop'
        ImportTitle          = 'Excel file'
        ImportHint           = "Вибрати xlsx"
        ExcelButton          = 'Вибрати Excel'
        PdfFolderLabel       = 'Вибір папки для зберігання паролів у PDF'
        PdfFolderButton      = 'Вибрати папку'
        SettingsTitle        = 'AD settings'
        DomainLabel          = 'Domain for UPN / mail'
        DomainTooltip        = 'Наприклад: donnu.edu.ua'
        PasswordNeverExpires = 'Password ніколи не закінчується'
        OuLabel              = 'Target OU'
        OuButton             = 'Вибрати OU'
        GroupsLabel          = 'Groups (SamAccountNames)'
        GroupsButton         = 'Вибрати Групи'
        GroupsTooltip        = "Введіть групи через кому або натисніть 'Вибрати групи'"
        RunButton            = 'Створити користувачів'
        LogTitle             = 'Execution log'
        ExcelErrorTitle      = 'Excel error'
        NoDataTitle          = 'No data'
        NoOuTitle            = 'No OU'
        NoDomainTitle        = 'No domain'
        NoPdfFolderTitle     = 'No PDF folder'
        PdfErrorTitle        = 'PDF error'
    }
}
