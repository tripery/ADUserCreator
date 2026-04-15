function Get-MainFormTexts {
    return [ordered]@{
        WindowTitle          = 'ADUserCreator Desktop'
        ImportTitle          = 'Excel файл'
        ImportHint           = "Джерело: *.xlsx з колонками 'Вступник' та 'Структурний підрозділ'"
        ExcelButton          = 'Вибрати Excel'
        PdfFolderLabel       = 'Папка для PDF з паролями'
        PdfFolderButton      = 'Вибрати папку'
        SettingsTitle        = 'Налаштування AD'
        DomainLabel          = 'Домен для UPN / пошти'
        DomainTooltip        = 'Наприклад: donnu.edu.ua'
        PasswordNeverExpires = 'Пароль не має терміну дії'
        OuLabel              = 'Цільовий OU'
        OuButton             = 'Вибрати OU'
        GroupsLabel          = 'Групи (SamAccountName через кому)'
        GroupsButton         = 'Вибрати групи'
        GroupsTooltip        = "Введіть групи через кому або натисніть 'Вибрати групи'"
        RunButton            = 'Створити користувачів'
        LogTitle             = 'Журнал виконання'
        ExcelErrorTitle      = 'Помилка Excel'
        NoDataTitle          = 'Немає даних'
        NoOuTitle            = 'Не вибрано OU'
        NoDomainTitle        = 'Не вказано домен'
        NoPdfFolderTitle     = 'Не вибрано папку для PDF'
        PdfErrorTitle        = 'Помилка PDF'
    }
}
