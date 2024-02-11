#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "debtoexcle_sence.h"
#include <QPainter>
#include <QPushButton>
#include <QAction>
#include <QDialog>
#include <QTimer>
#include <QMessageBox>
#include "exceltodbc_sence.h"
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    //设置固定大小
    setFixedSize(400,400);

    setWindowTitle("Excel_Dbc");
    setWindowIcon(QIcon(":/res/OIP (1).jpg"));

    QPushButton *ExcelToDbc_Button = new QPushButton("Excel转Dbc",this);
    ExcelToDbc_Button->move(this->width()/2-75,60);
    ExcelToDbc_Button->setFixedSize(150,80);

    QPushButton *DbcToExcel_Button = new QPushButton("Dbc转Excel",this);
    DbcToExcel_Button->move(this->width()/2-75,160);
    DbcToExcel_Button->setFixedSize(150,80);

    connect(ExcelToDbc_Button,&QPushButton::clicked,[=](){
        Exceltodbc_sence *exceltodbc_scene=new Exceltodbc_sence(this);
        exceltodbc_scene->show();

    });

    connect(DbcToExcel_Button,&QPushButton::clicked,[=](){
        DebToExcle_Sence *debToExcle_Sence=new DebToExcle_Sence(this);
        debToExcle_Sence->show();

    });


    connect(ui->actionactor,&QAction::triggered,[=](){
        QMessageBox *information=new QMessageBox(this);
        information->setText("作者：Pexlor\n邮箱：1018734423@qq.com");
        information->setWindowTitle("作者");
        information->show();
    });
}

MainWindow::~MainWindow()
{
    delete ui;
}

