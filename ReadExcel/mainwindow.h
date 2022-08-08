#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>
#include <QFile>
#include <QDebug>
#include <QFileDialog>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();



//   QAxObject *excel = NULL;    //本例中，excel设定为Excel文件的操作对象
//   QAxObject *workbooks = NULL;
//   QAxObject *workbook = NULL;  //Excel操作对象
//   int column_count;
//   QStringList data;
//   QString dbname;
//   QString sql_date;

//   int cellrow = 1;


private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
