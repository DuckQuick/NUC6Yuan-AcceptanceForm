#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QLabel>
#include <QStringList>

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
private slots:
    void slotLoadFile();
    // void slotLoadHistory();
    void slotCalculateFile();
    // void slotCalculateDetail();
    void slotWork();
    QString convertToChineseCurrency(double amount);
private:
    enum SubWindowState {
        Hidden,
        ShowingHistory,
        ShowingDetail
    };

    QLabel *qlCalculateFile;
    QLabel *qlLoadFile;

    QStringList qslLoadedFiles; // 缓存历史加载的文件名
    QStringList qslCalculateData;
    QList<QVariantList> qlInvoiceData;

};
#endif // MAINWINDOW_H
