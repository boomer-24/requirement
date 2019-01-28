#include "require.h"

Require::Require()
{
    this->wordApplication_ = new QAxObject("Word.Application");
    this->wordDocument_ = this->wordApplication_->querySubObject("Documents()");
    this->wordDocument_->querySubObject("Add()");
    this->activeDocument_ = this->wordApplication_->querySubObject("ActiveDocument()");

    this->selection_ = this->wordApplication_->querySubObject("Selection()");
    this->range_ = selection_->querySubObject("Range()");
    this->tables_ = this->activeDocument_->querySubObject("Tables()");
    this->font_ = selection_->querySubObject("Font");    

    this->mapMonth_.insert(1, "января");
    this->mapMonth_.insert(2, "февраля");
    this->mapMonth_.insert(3, "марта");
    this->mapMonth_.insert(4, "апреля");
    this->mapMonth_.insert(5, "мая");
    this->mapMonth_.insert(6, "июня");
    this->mapMonth_.insert(7, "июля");
    this->mapMonth_.insert(8, "августа");
    this->mapMonth_.insert(9, "сентября");
    this->mapMonth_.insert(10, "октября");
    this->mapMonth_.insert(11, "ноября");
    this->mapMonth_.insert(12, "декабря");
}

Require::Require(QString _name, QString _amount)
{
    this->wordApplication_ = new QAxObject("Word.Application");
    this->wordDocument_ = this->wordApplication_->querySubObject("Documents()");
    this->wordDocument_->querySubObject("Add()");
    this->activeDocument_ = this->wordApplication_->querySubObject("ActiveDocument()");

    this->selection_ = this->wordApplication_->querySubObject("Selection()");
    this->range_ = selection_->querySubObject("Range()");
    this->tables_ = this->activeDocument_->querySubObject("Tables()");
    this->font_ = selection_->querySubObject("Font");

    this->name_ = _name;
    this->amount_ = _amount;    

    this->mapMonth_.insert(1, "января");
    this->mapMonth_.insert(2, "февраля");
    this->mapMonth_.insert(3, "марта");
    this->mapMonth_.insert(4, "апреля");
    this->mapMonth_.insert(5, "мая");
    this->mapMonth_.insert(6, "июня");
    this->mapMonth_.insert(7, "июля");
    this->mapMonth_.insert(8, "августа");
    this->mapMonth_.insert(9, "сентября");
    this->mapMonth_.insert(10, "октября");
    this->mapMonth_.insert(11, "ноября");
    this->mapMonth_.insert(12, "декабря");
}

void Require::setNumber(int _number)
{
    this->number_ = _number;
}

void Require::setKarpunin(const QString &_karpunin)
{
    this->karpunin_ = _karpunin;
}

Require::~Require()
{
    this->wordApplication_->dynamicCall("Quit()");
    delete this->wordApplication_;
}

void Require::CreateList(QString _name, QString _amount)
{
    //    this->selection_->querySubObject("PageSetup()")->setProperty("TopMargin", 25);
    //    this->selection_->querySubObject("PageSetup()")->setProperty("BottomMargin", 25);    

    this->selection_->querySubObject("PageSetup()")->setProperty("LeftMargin", 22);
    this->selection_->querySubObject("PageSetup()")->setProperty("RightMargin", 22);
    this->font_->setProperty("Size", 12);
    this->font_->setProperty("Name", "Times New Roman");
    this->selection_->dynamicCall("TypeText(const QString&)", "Предприятие");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", "              ФГУП «НПЦ АП»               ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "                                                            ");
    this->font_->setProperty("Size", 10);
    this->selection_->dynamicCall("TypeText(const QString&)", "Типовая форма № М-11");

    //    this->selection_->dynamicCall("TypeParagraph()");
    //    this->selection_->dynamicCall("moveDown()");

    this->range_ = selection_->querySubObject("Range()");

    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 1, 2);
    this->font_->setProperty("Size", 14);
    this->selection_->dynamicCall("TypeText(const QString&)", "               Требование №");

    this->font_->setProperty("Underline", "wdUnderlineSingle");//////////////////////////////////////////////////////////
    this->font_->setProperty("Size", 12);

    this->selection_->dynamicCall("TypeText(const QString&)",
                                  QString("    ").append(QString::number(this->number_)).append("__"));
    this->font_->setProperty("Underline", "wdUnderlineNone");

    this->selection_->dynamicCall("TypeText(const QString&)", "\n                   ");
    this->selection_->dynamicCall("TypeText(const QString&)", "«");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", QDate::currentDate().day());
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "» ");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", this->mapMonth_.value(QDate::currentDate().month()));
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", QDate::currentDate().year());
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", " г.");

    //    this->selection_->dynamicCall("TypeText(const QString&)","__________\n"
    //                                                             "             «_____» ___________ _____ г.");

    QAxObject* cell_11 = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    cell_11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.6", "wdAdjustNone");
    cell_11->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    QAxObject* cell_12 = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    cell_12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "10.6", "wdAdjustNone");
    cell_12->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");


    this->selection_->dynamicCall("moveRight()");
    this->range_ = selection_->querySubObject("Range()");

    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 2, 4, 1, 2);

    this->font_->setProperty("Size", 11);
    QAxObject* cell_11_in = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    cell_11_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_11_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    QAxObject* rangeCell = cell_11_in->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Вид\nопер.");
    QAxObject* cell_21 = this->table_->querySubObject("Cell(Row, Column)", 2, 1);
    cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");

    QAxObject* cell_12_in = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    cell_12_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_12_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_12_in->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Склад");
    QAxObject* cell_22 = this->table_->querySubObject("Cell(Row, Column)", 2, 2);
    cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");

    QAxObject* cell_13 = this->table_->querySubObject("Cell(Row, Column)", 1, 3);
    cell_13->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_13->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_13->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Цех,\nотдел");
    QAxObject* cell_23 = this->table_->querySubObject("Cell(Row, Column)", 2, 3);
    cell_23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_23->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_23->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "138");

    QAxObject* cell_14_in = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    cell_14_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.6", "wdAdjustNone");
    cell_14_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    QAxObject* cell_24_in = this->table_->querySubObject("Cell(Row, Column)", 2, 4);
    cell_24_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.6", "wdAdjustNone");

    QAxObject* cell_14 = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    cell_14->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    cell_14->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_14->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Шифр затрат");
    QAxObject* cell_24 = this->table_->querySubObject("Cell(Row, Column)", 2, 4);
    cell_24->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    cell_24->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    rangeCell = cell_24->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "№ заказа");

    QAxObject* cell_34 = this->table_->querySubObject("Cell(Row, Column)", 3, 4);
    cell_34->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    cell_34->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_34->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "2323094");

//    QAxObject* cell_34 = this->table_->querySubObject("Cell(Row, Column)", 3, 4);
//    cell_34->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* cell_25 = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    cell_25->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_25->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "прибор, деталь");

    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");

    this->font_->setProperty("Size", 12);
    this->selection_->dynamicCall("TypeText(const QString&)", "\nЧерез кого ___");
    this->font_->setProperty("Underline", "wdUnderlineSingle");/////////////////////////////
    this->selection_->dynamicCall("TypeText(const QString&)", this->karpunin_);
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "_____ Затребовал ____________________ "
                                                              "Разрешил ____________________");
    this->range_ = selection_->querySubObject("Range()");

    this->font_->setProperty("Size", 8);
    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 4, 8, 1, 2);


    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 1);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.5", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 2);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.44", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 3);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.75", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 4);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.25", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_1 = this->table_->querySubObject("Cell(Row, Column)", i, 5);
        tcell_1->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.0", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 6);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 7);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.2", "wdAdjustNone");
    }

    for(int i = 1; i <= 5; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 8);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.0", "wdAdjustNone");
    }


    QAxObject* tcell_15 = this->table_->querySubObject("Cell(Row, Column)", 1, 5);
    tcell_15->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    QAxObject* tcell_25 = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    tcell_25->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* tcell_35 = this->table_->querySubObject("Cell(Row, Column)", 3, 5);
    tcell_35->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* tcell_45 = this->table_->querySubObject("Cell(Row, Column)", 4, 5);
    tcell_45->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    QAxObject* tcell_55 = this->table_->querySubObject("Cell(Row, Column)", 5, 5);
    tcell_55->dynamicCall("Split(NumRows, NumColumns)", 1, 2);


    //    QAxObject* rows = this->table_->querySubObject("Rows()");
    //    rows->dynamicCall("SetHeight(RowHeight, HeightRule)", 25, "wdRowHeightExactly");
    this->selection_->dynamicCall("moveRight()");
    QAxObject* allign = this->selection_->querySubObject("ParagraphFormat()");

    //    qDebug() << this->table_->querySubObject("Columns()")->dynamicCall("Count").toString();

    QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Номенклатурный\n        номер");
    allign->dynamicCall("Alignment", "wdAlignParagraphCenter");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Вид\nприемки");
    this->selection_->dynamicCall("moveRight()");
    allign->dynamicCall("Alignment", "wdAlignParagraphCenter");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 3);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Наименование, сорт, размер");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 3);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", _name);
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    //    tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Ед.\nизм.");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 4);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "шт");
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "          Количество");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "затребовано");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", QString("       ").append(_amount));
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 2, 6);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "отпущено");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 6);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Цена");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 7);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Сумма");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 8);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "   Поряд. №\n     записи\n      по скл.\n       карт.");

    for (int i = 0; i < 10; i ++)
        this->selection_->dynamicCall("moveDown()");
    this->font_->setProperty("Size", 12);
    this->selection_->dynamicCall("TypeText(const QString&)", "Изделие:\n");
    this->selection_->dynamicCall("TypeText(const QString&)", "                    "
                                                              "Отпустил    ____________________	   "
                                                              "Получил ____________________");

    for (int i = 0; i < 5; i++)
        this->selection_->dynamicCall("TypeParagraph()");

    this->selection_->querySubObject("InlineShapes")->dynamicCall("AddHorizontalLineStandard()");



    for (int i = 0; i < 2; i++)
        this->selection_->dynamicCall("TypeParagraph()");

    this->font_->setProperty("Size", 12);
    this->font_->setProperty("Name", "Times New Roman");
    this->selection_->dynamicCall("TypeText(const QString&)", "Предприятие");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", "              ФГУП «НПЦ АП»               ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "                                                            ");
    this->font_->setProperty("Size", 10);
    this->selection_->dynamicCall("TypeText(const QString&)", "Типовая форма № М-11");

    //    this->selection_->dynamicCall("TypeParagraph()");
    //    this->selection_->dynamicCall("moveDown()");

    this->range_ = selection_->querySubObject("Range()");

    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 1, 2);
    this->font_->setProperty("Size", 14);
    this->selection_->dynamicCall("TypeText(const QString&)", "               Требование №");

    this->font_->setProperty("Underline", "wdUnderlineSingle");//////////////////////////////////////////////////////////
    this->font_->setProperty("Size", 12);

    this->selection_->dynamicCall("TypeText(const QString&)",
                                  QString("    ").append(QString::number(this->number_)).append("__"));
    this->font_->setProperty("Underline", "wdUnderlineNone");

    this->selection_->dynamicCall("TypeText(const QString&)", "\n                   ");
    this->selection_->dynamicCall("TypeText(const QString&)", "«");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", QDate::currentDate().day());
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "» ");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", this->mapMonth_.value(QDate::currentDate().month()));
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineSingle");
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->selection_->dynamicCall("TypeText(const QString&)", QDate::currentDate().year());
    this->selection_->dynamicCall("TypeText(const QString&)", " ");
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", " г.");

    //    this->selection_->dynamicCall("TypeText(const QString&)","__________\n"
    //                                                             "             «_____» ___________ _____ г.");

    cell_11 = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    cell_11->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "8.6", "wdAdjustNone");
    cell_11->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    cell_12 = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    cell_12->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "10.6", "wdAdjustNone");
    cell_12->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");


    this->selection_->dynamicCall("moveRight()");
    this->range_ = selection_->querySubObject("Range()");

    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 2, 4, 1, 2);

    this->font_->setProperty("Size", 11);
    cell_11_in = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    cell_11_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_11_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_11_in->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Вид\nопер.");
    cell_21 = this->table_->querySubObject("Cell(Row, Column)", 2, 1);
    cell_21->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");

    cell_12_in = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    cell_12_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_12_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_12_in->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Склад");
    cell_22 = this->table_->querySubObject("Cell(Row, Column)", 2, 2);
    cell_22->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");

    cell_13 = this->table_->querySubObject("Cell(Row, Column)", 1, 3);
    cell_13->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_13->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_13->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Цех,\nотдел");
    cell_23 = this->table_->querySubObject("Cell(Row, Column)", 2, 3);
    cell_23->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    cell_23->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_23->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "138");

    cell_14_in = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    cell_14_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.6", "wdAdjustNone");
    cell_14_in->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    cell_24_in = this->table_->querySubObject("Cell(Row, Column)", 2, 4);
    cell_24_in->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.6", "wdAdjustNone");

    cell_14 = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    cell_14->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    cell_14->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_14->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Шифр затрат");
    cell_24 = this->table_->querySubObject("Cell(Row, Column)", 2, 4);
    cell_24->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    cell_24->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    rangeCell = cell_24->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "№ заказа");

    cell_34 = this->table_->querySubObject("Cell(Row, Column)", 3, 4);
    cell_34->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    cell_34->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_34->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "2323094");
    cell_25 = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    cell_25->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = cell_25->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "прибор, деталь");

    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");
    this->selection_->dynamicCall("moveDown()");

    this->font_->setProperty("Size", 12);    this->selection_->dynamicCall("TypeText(const QString&)", "\nЧерез кого ___");
    this->font_->setProperty("Underline", "wdUnderlineSingle");/////////////////////////////
    this->selection_->dynamicCall("TypeText(const QString&)", this->karpunin_);
    this->font_->setProperty("Underline", "wdUnderlineNone");
    this->selection_->dynamicCall("TypeText(const QString&)", "_____ Затребовал ____________________ "
                                                              "Разрешил ____________________");
    this->range_ = selection_->querySubObject("Range()");

    this->font_->setProperty("Size", 8);
    this->table_ = tables_->querySubObject("Add(Range,NumRows,NumColumns, DefaulttableBehavior, AutoFitBehavior)",
                                           range_->asVariant(), 4, 8, 1, 2);


    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 1);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "2.5", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 2);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.44", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 3);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "5.75", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 4);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.25", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_1 = this->table_->querySubObject("Cell(Row, Column)", i, 5);
        tcell_1->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.0", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 6);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    }

    for(int i = 1; i <= 4; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 7);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.2", "wdAdjustNone");
    }

    for(int i = 1; i <= 5; i++)
    {
        QAxObject* tcell_ = this->table_->querySubObject("Cell(Row, Column)", i, 8);
        tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "3.0", "wdAdjustNone");
    }


    tcell_15 = this->table_->querySubObject("Cell(Row, Column)", 1, 5);
    tcell_15->dynamicCall("Split(NumRows, NumColumns)", 2, 1);
    tcell_25 = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    tcell_25->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    tcell_35 = this->table_->querySubObject("Cell(Row, Column)", 3, 5);
    tcell_35->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    tcell_45 = this->table_->querySubObject("Cell(Row, Column)", 4, 5);
    tcell_45->dynamicCall("Split(NumRows, NumColumns)", 1, 2);
    tcell_55 = this->table_->querySubObject("Cell(Row, Column)", 5, 5);
    tcell_55->dynamicCall("Split(NumRows, NumColumns)", 1, 2);


    //    QAxObject* rows = this->table_->querySubObject("Rows()");
    //    rows->dynamicCall("SetHeight(RowHeight, HeightRule)", 25, "wdRowHeightExactly");
    this->selection_->dynamicCall("moveRight()");
    allign = this->selection_->querySubObject("ParagraphFormat()");

    //    qDebug() << this->table_->querySubObject("Columns()")->dynamicCall("Count").toString();

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 1);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Номенклатурный\n        номер");
    allign->dynamicCall("Alignment", "wdAlignParagraphCenter");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 2);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Вид\nприемки");
    this->selection_->dynamicCall("moveRight()");
    allign->dynamicCall("Alignment", "wdAlignParagraphCenter");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 3);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Наименование, сорт, размер");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 3);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", _name);
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 4);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    //    tcell_->dynamicCall("SetWidth(ColumnWidth, RulerStyle)", "1.5", "wdAdjustNone");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Ед.\nизм.");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 4);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "шт");
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "          Количество");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 2, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "затребовано");

    this->font_->setProperty("Size", 10);
    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 3, 5);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", QString("       ").append(_amount));
    this->font_->setProperty("Size", 8);

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 2, 6);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "отпущено");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 6);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Цена");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 7);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "Сумма");

    tcell_ = this->table_->querySubObject("Cell(Row, Column)", 1, 8);
    tcell_->setProperty("VerticalAlignment", "wdCellAlignVerticalCenter");
    rangeCell = tcell_->querySubObject("Range()");
    rangeCell->dynamicCall("InsertAfter(Text)", "   Поряд. №\n     записи\n      по скл.\n       карт.");

    for (int i = 0; i < 10; i ++)
        this->selection_->dynamicCall("moveDown()");
    this->font_->setProperty("Size", 12);
    this->selection_->dynamicCall("TypeText(const QString&)", "Изделие:\n");
    this->selection_->dynamicCall("TypeText(const QString&)", "                    "
                                                              "Отпустил    ____________________	   "
                                                              "Получил ____________________");

//    this->number_++;
}

void Require::NewList()
{
    this->selection_->dynamicCall("InsertBreak()", "wdPageBreak"); //новый лист
}

void Require::SaveList()
{
    this->activeDocument_->dynamicCall("SaveAs(FileName)", "C:/D/1.docx");
    this->activeDocument_->dynamicCall("Close()");
}

void Require::SaveList(QString _dirPath)
{
    this->activeDocument_->dynamicCall("SaveAs(FileName)", _dirPath.append(".docx"));
    this->activeDocument_->dynamicCall("Close()");
}
