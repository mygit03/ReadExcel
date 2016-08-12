#ifndef READEXCEL_H
#define READEXCEL_H

#include <QMainWindow>

namespace Ui {
class ReadExcel;
}

class ReadExcel : public QMainWindow
{
    Q_OBJECT

public:
    explicit ReadExcel(QWidget *parent = 0);
    ~ReadExcel();

    void init();                                //初始化
    void readToTable(const QString sheet);      //读Excel到tableWidget
    void writeToExcel(const QString sheet);     //保存修改内容到Excel

private slots:
    void slot_openExcel();

    void slot_saveExcel();

    void on_comboBox_activated(const QString &arg1);

    void on_tableWidget_cellChanged(int row, int column);

private:
    Ui::ReadExcel *ui;

    QString m_fileName;       //文件名
    QList<QPoint> cellList; //记录修改内容的位置
};

#endif // READEXCEL_H
