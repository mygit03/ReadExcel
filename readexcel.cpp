#include "readexcel.h"
#include "ui_readexcel.h"

#include "about.h"

#include <QFileDialog>
#include <QtXlsx>
#include <QDebug>
#include <QStandardPaths>
#include <QTableWidgetItem>
#include <QMessageBox>

ReadExcel::ReadExcel(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::ReadExcel)
{
    ui->setupUi(this);

    init();
}

ReadExcel::~ReadExcel()
{
    delete ui;
}

void ReadExcel::init()
{
    setWindowTitle(tr("读Excel文件"));

    ui->tableWidget->horizontalHeader()->setVisible(false);
    ui->tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    connect(ui->btn_openExcel, SIGNAL(clicked()), this, SLOT(slot_openExcel()));
    connect(ui->btn_saveExcel, SIGNAL(clicked()), this, SLOT(slot_saveExcel()));
    connect(ui->act_open, SIGNAL(triggered()), this, SLOT(slot_openExcel()));
    connect(ui->act_save, SIGNAL(triggered()), this, SLOT(slot_saveExcel()));
    connect(ui->act_exit, SIGNAL(triggered()), this, SLOT(close()));
    connect(ui->act_aboutQt, SIGNAL(triggered()), qApp, SLOT(aboutQt()));
}

void ReadExcel::readToTable(const QString sheet)
{
    QXlsx::Document xlsx(m_fileName);
    xlsx.selectSheet(sheet);                        //设置当前Sheet

    //获取当前Sheet的行列
    QXlsx::CellRange range = xlsx.dimension();
    int rows = range.rowCount();
    int cols = range.columnCount();

    //清空并设置tableWidget的行列
    ui->tableWidget->clear();
    ui->tableWidget->setRowCount(rows);
    ui->tableWidget->setColumnCount(cols);
    qDebug() << "curSheet:" << xlsx.currentSheet()->sheetName() << sheet;
    for(int i = 0; i < rows; i++){
        for (int j = 0; j < cols; j++){
            QString text = xlsx.read(i+1, j+1).toString();
//            qDebug() << "text:" << text;
            QTableWidgetItem *item = new QTableWidgetItem(text);
            item->setTextAlignment(Qt::AlignCenter);                //设置居中对齐
            ui->tableWidget->setItem(i, j, item);
        }
    }

    cellList.clear();       //读完Excel清空修改记录
    qDebug() << "cellList:" << cellList;
}

void ReadExcel::writeToExcel(const QString sheet)
{
    qDebug() << "write:" << sheet;
    QXlsx::Document xlsx(m_fileName);
    xlsx.selectSheet(sheet);

    //遍历修改记录并写入Excel文件保存
    foreach (QPoint point, cellList) {
        int row = point.x();
        int col = point.y();
        QString text = ui->tableWidget->item(row - 1, col - 1)->text();
        xlsx.write(row, col, text);
        qDebug() << "write:" << point << row << col << text;
    }
    xlsx.save();
}

void ReadExcel::slot_openExcel()
{
    QString path = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
    QString fileName = QFileDialog::getOpenFileName(this, tr("打开Excel文件"), path,
                                            tr("Excel(*.xls *.xlsx);;All File(*.*)"));
    qDebug() << "fileName:" << fileName;
    if(fileName.isEmpty()){
        return;
    }
    m_fileName = fileName;
    ui->lineEdit->setText(m_fileName);
//    xlsx.setObjectName(fileName);
//    xlsx.saveAs("fileName.xls");
    QXlsx::Document xlsx(m_fileName);
    QStringList sheetList = xlsx.sheetNames();      //获取Excel文件的所有表单
    qDebug() << "sheetList:" << sheetList << QFileInfo(m_fileName).fileName();
    ui->comboBox->addItems(sheetList);

    cBox_text = ui->comboBox->currentText();
    qDebug() << "cBox_text open:" << cBox_text;

    QString curSheet = ui->comboBox->currentText();
    readToTable(curSheet);

    ui->statusBar->showMessage(tr("文件已导入!"), 3000);
}

void ReadExcel::slot_saveExcel()
{
    qDebug() << "save!" << cBox_text << ui->comboBox->currentText();
    writeToExcel(cBox_text);
    ui->statusBar->showMessage(tr("文件已保存!"), 3000);
}

void ReadExcel::on_comboBox_activated(const QString &arg1)
{
    qDebug() << "arg1:" << arg1;
    if (!cellList.isEmpty()){
        int ret = QMessageBox::warning(this, tr("警告"), tr("文件已修改,是否保存?"),
                                       tr("是"), tr("否"), 0, 1);
        switch (ret) {
        case 0:
            qDebug() << "cBox_text:::" << cBox_text;
            writeToExcel(cBox_text);
            break;
        default:
            break;
        }
    }
    cBox_text = arg1;
    qDebug() << "cBox_text:" << cBox_text;
    readToTable(cBox_text);
}

void ReadExcel::on_tableWidget_cellChanged(int row, int column)
{
//    qDebug() << "cellChanged:" << row << column;
    QPoint point(row + 1, column + 1);
    cellList.append(point);
//    qDebug() << "point:" << point << point.x() << point.y() << "cellList:" << cellList;
}

void ReadExcel::on_act_about_triggered()
{
    About dlg;
    dlg.exec();
}
