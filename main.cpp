#include "readexcel.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    ReadExcel w;
    w.show();

    return a.exec();
}
