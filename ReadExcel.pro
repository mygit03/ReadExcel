#-------------------------------------------------
#
# Project created by QtCreator 2016-08-12T18:54:23
#
#-------------------------------------------------

QT       += core gui xlsx

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ReadExcel
TEMPLATE = app


SOURCES += main.cpp\
        readexcel.cpp

HEADERS  += readexcel.h

FORMS    += readexcel.ui

DISTFILES += \
    MyApp.rc
RC_FILE += \
    MyApp.rc

RESOURCES += \
    images.qrc
