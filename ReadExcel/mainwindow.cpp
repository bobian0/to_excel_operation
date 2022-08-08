#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    

    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", true);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open(const QString&)", "C:\\Users\\Administrator\\Desktop\\10_17_16.xlsx");
    //excel.setProperty("Caption", "Qt Excel");
    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("WorkSheets");  //Sheets也可换用WorkSheets

    //删除工作表（删除第一个）
    QAxObject *first_sheet = work_sheets->querySubObject("Item(int)", 1);
    first_sheet->dynamicCall("delete");

    //插入工作表（插入至最后一行）
    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目
    QAxObject *last_sheet = work_sheets->querySubObject("Item(int)", sheet_count);
    QAxObject *work_sheet = work_sheets->querySubObject("Add(QVariant)", last_sheet->asVariant());
    last_sheet->dynamicCall("Move(QVariant)", work_sheet->asVariant());

    //work_sheet->setProperty("Name", "Qt Sheet");  //设置工作表名称

    //操作单元格（第2行第2列）
//    QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", 2, 2);
//    cell->setProperty("Value", "Java C++ C# PHP Perl Python Delphi Ruby");  //设置单元格值
//    cell->setProperty("RowHeight", 50);  //设置单元格行高
//    cell->setProperty("ColumnWidth", 30);  //设置单元格列宽
//    cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
//    cell->setProperty("VerticalAlignment", -4108);  //上对齐（xlTop）-4160 居中（xlCenter）：-4108  下对齐（xlBottom）：-4107
//    cell->setProperty("WrapText", true);  //内容过多，自动换行
//    //cell->dynamicCall("ClearContents()");  //清空单元格内容

//    QAxObject* interior = cell->querySubObject("Interior");
//    interior->setProperty("Color", QColor(0, 255, 0));   //设置单元格背景色（绿色）

//    QAxObject* border = cell->querySubObject("Borders");
//    border->setProperty("Color", QColor(0, 0, 255));   //设置单元格边框色（蓝色）

//    QAxObject *font = cell->querySubObject("Font");  //获取单元格字体
//    font->setProperty("Name", QStringLiteral("华文彩云"));  //设置单元格字体
//    font->setProperty("Bold", true);  //设置单元格字体加粗
//    font->setProperty("Size", 20);  //设置单元格字体大小
//    font->setProperty("Italic", true);  //设置单元格字体斜体
//    font->setProperty("Underline", 2);  //设置单元格下划线
//    font->setProperty("Color", QColor(255, 0, 0));  //设置单元格字体颜色（红色）
    int i = 10,k = 10;
    //设置单元格内容，并合并单元格（第5行第3列-第8行第5列）
    QAxObject *cell_5_6 = work_sheet->querySubObject("Cells(int,int)", i, k);
    cell_5_6->setProperty("Value", "Java");  //设置单元格值
    QAxObject *cell_8_5 = work_sheet->querySubObject("Cells(int,int)", 8, 5);
    cell_8_5->setProperty("Value", "C++");

//    QString merge_cell;
//    merge_cell.append(QChar(3 - 1 + 'A'));  //初始列
//    merge_cell.append(QString::number(5));  //初始行
//    merge_cell.append(":");
//    merge_cell.append(QChar(5 - 1 + 'A'));  //终止列
//    merge_cell.append(QString::number(8));  //终止行
//    QAxObject *merge_range = work_sheet->querySubObject("Range(const QString&)", merge_cell);
//    merge_range->setProperty("HorizontalAlignment", -4108);
//    merge_range->setProperty("VerticalAlignment", -4108);
//    merge_range->setProperty("WrapText", true);
//    merge_range->setProperty("MergeCells", true);  //合并单元格
//    //merge_range->setProperty("MergeCells", false);  //拆分单元格

    work_book->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
    work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
    excel.dynamicCall("Quit(void)");  //退出













//    excel = new QAxObject("Excel.Application");
//    excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
//    excel->setProperty("EnableEvents",false);
//    workbooks = excel->querySubObject("WorkBooks");

//    QString nums = QFileDialog::getOpenFileName(this,tr("选择文件"),"./",tr("*.xls *.xlsx"));
//    if(nums.isEmpty())
//    {
//        return;
//    }
//    QFile *file = new QFile;
//    file->setFileName(nums);
//    if(file->open(QIODevice::ReadOnly))
//    {
//        workbook = workbooks->querySubObject("Open(QString&)", nums);
//        // 获取打开的excel文件中所有的工作sheet
//        workbook = excel->querySubObject("ActiveWorkBook");  //获取工作簿
//        QAxObject * worksheets = workbook->querySubObject("WorkSheets");
//        int iWorkSheet = worksheets->property("Count").toInt();  //获取工作表的数目

//        //qDebug() << QString("Excel文件中表的个数: %1").arg(QString::number(iWorkSheet));
//        QAxObject* pWorkSheet = workbook->querySubObject("Sheets(int)", 1);//获取第一张表
//        QAxObject* used_range = pWorkSheet->querySubObject("UsedRange"); //获取该sheet的使用范围对象

//        QVariant var = used_range->dynamicCall("Value");

//        QAxObject *rows  = used_range->querySubObject("Rows");
//        QAxObject *columns = used_range->querySubObject("Columns");

//        int row_start = used_range->property("Row").toInt();          //获得开始行
//        int column_start  = used_range->property("Column").toInt();     //获得开始列
//        int row_count = rows->property("Count").toInt();
//        column_count = columns->property("Count").toInt();

//        delete used_range;
//        QVariantList varRows = var.toList();            //得到表格中的所有数据
//        if(varRows.isEmpty()){return;}
//        const int rowCount = varRows.size();
//        qDebug()<<"总行数："<<row_count;
//        qDebug()<<"总列数："<<column_count;


//        for(int i = 0;i<rowCount-1;i++)
//        {
//            QVariantList rowData = varRows[i].toList();
//            qDebug()<<"数据:"<<i<<rowData[4].toString();
//        }

//        QAxObject * range = pWorkSheet->querySubObject("Cells(int,int)", 1, 1 );


//        workbook->dynamicCall("Close()");
//        excel->dynamicCall("Quit()");       //断开连接，接收新的连接
//        file->close();                      //关闭文件
//    }
//    else
//    {
//        qDebug()<<"没有文件";
//    }
//    workbook->dynamicCall("Close (Boolean)", false);  //关闭文件
//    delete excel;               //回收指针
//    excel = NULL;


}

MainWindow::~MainWindow()
{
    delete ui;
}

