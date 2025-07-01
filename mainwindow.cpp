#include "mainwindow.h"
#include <QPushButton>
#include <QWidget>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QLabel>
#include <QUrl>
#include <QGridLayout>
#include <QFileDialog>
#include <QMessageBox>
#include <QAxObject>
#include <QDebug>
#include <QFileInfo>
#include <QList>
#include <QChar>
#include <QCoreApplication>
#include <QFile>
#include <QStandardPaths>
#include <QApplication>
#include <QDesktopServices>
#include <QIcon>
#include <QSystemTrayIcon>
#include <QMenu>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    QApplication::setStyle("windowsvista");
    // resize(360, 240);
    setFixedSize(360, 240);
    setWindowTitle("验收单一键填写工具 - 立创商城版");
    setWindowIcon(QIcon(":/cantaloupe.ico"));

    QSystemTrayIcon* pSystemTray = new QSystemTrayIcon();
    if (pSystemTray != NULL) {
        pSystemTray->setIcon(QIcon(":/cantaloupe.ico"));
        pSystemTray->setToolTip("验收单一键填写工具");
        QAction *exitAction = new QAction("退出", this);
        QMenu *trayMenu = new QMenu(this);
        connect(exitAction, &QAction::triggered, qApp, &QApplication::quit);
        trayMenu->addAction(exitAction);
        pSystemTray->show();
    }


    qlInvoiceData.clear(); // 清空之前的数据

    auto lTitle = new QLabel("验收单一键填写工具", this);
    lTitle->setAlignment(Qt::AlignCenter);
    lTitle->setStyleSheet(
        "QLabel {"
        "    font-size: 20pt;"
        "    font-family: 'Verdana';"
        "    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #FFD700, stop:1 #FF8C00);"
        "    color: black;"
        "    border: 1px solid #333;"
        "    padding: 2px;"
        "    border-radius: 5px;"
        "    background-clip: text;"
        "}");

    auto lTip = new QLabel;
    lTip->setText("<html><head/><body>"
                  "<p><font color='red'><strong>注意：</strong></font>加载文件为"
                  "<a href=\"http://szlcsc.com\">立创商城</a>下的 "
                  "<a href=\"http://order.szlcsc.com/member/invoice/schedule.html\">"
                  "会员中心>交易管理>我的发票></a></p>"
                  "<p>"
                  "点击\"导出开票明细\"按钮后下载的.xls文件"
                  "</p>"
                  "</body></html");
    lTip->setTextInteractionFlags(Qt::TextBrowserInteraction);
    connect(lTip, &QLabel::linkActivated, [](const QString &link) {
        QDesktopServices::openUrl(QUrl(link));
    });


    auto btnLoadFile = new QPushButton("加载文件", this);
    connect(btnLoadFile, &QPushButton::clicked, this, &MainWindow::slotLoadFile);

    qlLoadFile = new QLabel(this);
    qlLoadFile->setText("请加载.xls文件");

    auto btnCalculateFile = new QPushButton("计算数据", this);
    connect(btnCalculateFile, &QPushButton::clicked, this, &MainWindow::slotCalculateFile);

    qlCalculateFile = new QLabel;
    qlCalculateFile->setText("总共 0 元");

    auto btnWork = new QPushButton("生成文件", this);
    btnWork->setFixedHeight(40);
    connect(btnWork, &QPushButton::clicked, this, &MainWindow::slotWork);

    auto girdMainLayout = new QGridLayout();
    girdMainLayout->setContentsMargins(30,0,30,0);
    girdMainLayout->addWidget(btnLoadFile, 0, 0, 1, 1);
    girdMainLayout->addWidget(qlLoadFile, 0, 1, 1, 3);
    girdMainLayout->addWidget(btnCalculateFile, 1, 0, 1, 1);
    girdMainLayout->addWidget(qlCalculateFile, 1, 1, 1, 3);
    girdMainLayout->addWidget(btnWork, 2, 1, 2, 2);

    auto vboxMainLayout = new QVBoxLayout;

    vboxMainLayout->addWidget(lTitle);
    vboxMainLayout->addWidget(lTip);
    vboxMainLayout->addSpacing(5);
    vboxMainLayout->addLayout(girdMainLayout);
    vboxMainLayout->addStretch();

    auto mainWidget = new QWidget(this);
    mainWidget->setLayout(vboxMainLayout);

    setCentralWidget(mainWidget);

}

void MainWindow::slotLoadFile()
{
    QString filePath = QFileDialog::getOpenFileName(this, "选择 Excel 文件", "", "Excel 文件 (*.xls *.xlsx)");
    if (filePath.isEmpty()) {
        QMessageBox::warning(this, "未选择文件", "请先选择一个文件！");
        return;
    }
    try {
        QAxObject excel("Excel.Application");
        if (excel.isNull()) {
            throw std::runtime_error("无法启动Excel实例。");
        }
        excel.setProperty("Visible", false);

        QAxObject *workbooks = excel.querySubObject("Workbooks");
        QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", filePath);
        QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);
        QAxObject *usedRange = worksheet->querySubObject("UsedRange");
        QAxObject *rows = usedRange->querySubObject("Rows");
        int lastRow = rows->property("Count").toInt();
        QString primaryCol = "H";
        QString primaryRange = QString("%1%2:%1%3").arg(primaryCol).arg(3).arg(lastRow); // 从第3行开始到最后一行

        QAxObject *range = worksheet->querySubObject("Range(const QString&)", primaryRange);
        if (!range || range->isNull()) {
            throw std::runtime_error("指定范围无效。");
        }

        QVariant primaryVar = range->property("Value");
        QStringList qslCacheData;
        if (primaryVar.isValid()) {
            QVariantList primaryRows = primaryVar.toList();
            for (const QVariant &innerVar : primaryRows) {
                QVariantList innerList = innerVar.toList();
                for (const QVariant &cellVar : innerList) {
                    QString cellValue = cellVar.toString().trimmed();
                    if (!cellValue.isEmpty()) {
                        qslCacheData.append(cellValue); // 保存商品名称到 qslCacheData
                    }
                }
            }
        }

        int validRowCount = qslCacheData.size(); // 商品名称列的有效行数

        QStringList columns = {"I", "L", "J", "N", "O"};
        for (const QString &col : columns) {
            QString colRange = QString("%1%2:%1%3").arg(col).arg(3).arg(2 + validRowCount);
            QAxObject *colRangeObj = worksheet->querySubObject("Range(const QString&)", colRange);
            if (!colRangeObj || colRangeObj->isNull()) {
                qDebug() << "Failed to get range for column:" << col;
                continue;
            }

            QVariant colVar = colRangeObj->property("Value");
            if (!colVar.isValid() || colVar.isNull()) {
                continue;
            }
            QVariantList outerList = colVar.toList();
            for (const QVariant &innerVar : outerList) {
                QVariantList innerList = innerVar.toList();
                for (const QVariant &cellVar : innerList) {
                    QString cellValue = cellVar.toString().trimmed();
                    qslCacheData.append(cellValue.isEmpty() ? "" : cellValue);
                }
            }
        }

        // 遍历每行并加入 qlInvoiceData
        for (int i = 0; i < validRowCount; ++i) {
            QString productName = (i < qslCacheData.size()) ? qslCacheData.at(i) : "";
            QString model = ((i + validRowCount) < qslCacheData.size()) ? qslCacheData.at(i + validRowCount) : "";
            QString quantity = ((i + 2 * validRowCount) < qslCacheData.size()) ? qslCacheData.at(i + 2 * validRowCount) : "0";
            QString unit = ((i + 3 * validRowCount) < qslCacheData.size()) ? qslCacheData.at(i + 3 * validRowCount) : "";
            QString totalAmount = ((i + 4 * validRowCount) < qslCacheData.size()) ? qslCacheData.at(i + 4 * validRowCount) : "0.0";
            QString discount = ((i + 5 * validRowCount) < qslCacheData.size()) ? qslCacheData.at(i + 5 * validRowCount) : "0.0";

            QVariantList rowData = {productName, model, quantity, unit, totalAmount, discount};
            qlInvoiceData.append(rowData);
        }
        // Debug 输出结果
        qDebug() << qlInvoiceData;

        qslLoadedFiles.append(filePath);
        QString fileName = QFileInfo(filePath).fileName();
        qlLoadFile->setText(fileName);
        workbook->dynamicCall("Close()");
        worksheet = nullptr;
        usedRange = nullptr;
        rows = nullptr;
        excel.dynamicCall("Quit()");

        QMessageBox::information(this, "加载成功", QString("有效行数：%1\n文件内容已缓存！").arg(validRowCount));
        qDebug() << "5";
    } catch (const std::exception &e) {
        QMessageBox::critical(this, "加载失败", e.what());
    }
}


void MainWindow::slotCalculateFile()
{
    if (qlInvoiceData.isEmpty()) {
        qlCalculateFile->setText("总共 0 元");
        return;
    }
    double totalSum = 0.0;
    for (const QVariantList &row : qlInvoiceData) {
        if (row.size() < 5) {
            continue; // 跳过无效行
        }
        // 获取 totalAmount 和 discount
        bool ok1 = false, ok2 = false;
        double totalAmount = row[4].toString().toDouble(&ok1); // 第 5 列 (totalAmount)
        double discount = row[5].toString().toDouble(&ok2);    // 第 6 列 (discount)
        if (ok1 && ok2) {
            totalSum += (totalAmount - discount);
        }
    }
    qlCalculateFile->setText(QString("总共 %1 元").arg(totalSum, 0, 'f', 2)); // 保留两位小数
}

void MainWindow::slotWork()
{
    try {
        if (qlInvoiceData.isEmpty()) {
            QMessageBox::warning(this, "无数据", "请先加载数据！");
            return;
        }

        int maxRowsPerFile = 17; // 每个文件最多17行数据
        int totalRows = qlInvoiceData.size();
        int fileCount = (totalRows + maxRowsPerFile - 1) / maxRowsPerFile; // 计算需要的文件数量

        QString rootDir = QCoreApplication::applicationDirPath(); // 软件所在根目录
        QString tempDirPath = rootDir + "/temp";
        QDir tempDir(tempDirPath);

        if (!tempDir.exists()) {
            tempDir.mkpath("."); // 创建临时文件夹
        }

        QString templatePath = ":/SourceForm.xlsx"; // 模板文件路径

        QAxObject excel("Excel.Application");
        if (excel.isNull()) {
            throw std::runtime_error("无法启动Excel实例。");
        }
        excel.setProperty("Visible", false); // 设置Excel不可见
        excel.setProperty("DisplayAlerts", false); // 关闭提示

        for (int fileIndex = 0; fileIndex < fileCount; ++fileIndex) {
            QString tempFileName = QString("AcceptanceFormTemp%1.xlsx").arg(fileIndex + 1);
            QString tempFilePath = tempDirPath + "/" + tempFileName;
            QFile::copy(templatePath, tempFilePath); // 复制模板文件

            QAxObject *workbooks = excel.querySubObject("Workbooks");
            QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", tempFilePath);
            QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

            int startRow = fileIndex * maxRowsPerFile;
            int endRow = std::min(startRow + maxRowsPerFile, totalRows);

            for (int i = startRow; i < endRow; ++i) {
                QVariantList row = qlInvoiceData[i];
                if (row.size() < 6) continue;

                QString productName = row[0].toString();
                QString model = row[1].toString();
                int quantity = row[2].toInt();
                QString unit = row[3].toString();
                double totalAmount = row[4].toDouble();
                double discount = row[5].toDouble();
                double unitPrice = quantity > 0 ? (totalAmount - discount) / quantity : 0.01;

                unitPrice = qMax(qRound(unitPrice * 100) / 100.0, 0.01);

                int excelRow = 5 + (i - startRow);
                QString startCell = QString("E%1").arg(excelRow);
                QString endCell = QString("I%1").arg(excelRow);
                QAxObject *range = worksheet->querySubObject("Range(const QString&)", startCell + ":" + endCell);
                QVariantList rowData = {productName, model, quantity, unit, unitPrice};
                range->setProperty("Value", QVariant(rowData));

                QAxObject *nameRange = worksheet->querySubObject("Range(const QString&)", QString("E%1:F%1").arg(excelRow));
                QAxObject *otherRange = worksheet->querySubObject("Range(const QString&)", QString("G%1:I%1").arg(excelRow));
                nameRange->setProperty("HorizontalAlignment", -4131); // 左对齐
                otherRange->setProperty("HorizontalAlignment", -4108); // 居中对齐

                double netAmount = totalAmount - discount;
                QStringList digits = QString::number(netAmount, 'f', 2).split('.');
                QString integerPart = digits[0];
                QString fractionPart = digits.size() > 1 ? digits[1] : "00";

                QList<QChar> intDigits;
                for (const QChar &ch : integerPart.rightJustified(4, '0').mid(0, 4)) {
                    intDigits.append(ch);
                }

                QList<QChar> fracDigits;
                for (const QChar &ch : fractionPart.leftJustified(2, '0')) {
                    fracDigits.append(ch);
                }

                QStringList columnLetters = {"M", "N", "O", "P", "Q", "R"};
                QList<QChar> amountDigits = intDigits + fracDigits;

                if (amountDigits.size() < columnLetters.size()) {
                    amountDigits.append(QChar('0'));
                }
                int coln = amountDigits.size();
                for (int col = 0; col < columnLetters.size(); ++col) {
                    if (col == 0 && amountDigits[col] == '0') {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", "");
                        coln = coln - 1;
                    } else if (col == 1 && amountDigits[col] == '0' && coln == amountDigits.size() - 1) {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", "");
                        coln = coln - 1;
                    }
                    else if (col == 2 && amountDigits[col] == '0' && coln == amountDigits.size() - 2) {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", "");
                        coln = coln - 1;
                    }
                    else if (col == 3 && amountDigits[col] == '0' && coln == amountDigits.size() - 3) {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", "");
                        coln = coln - 1;
                    }
                    else if (col == 4 && amountDigits[col] == '0' && coln == amountDigits.size() - 4) {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", "");
                        coln = coln - 1;
                    }
                    else {
                        QString cell = QString("%1%2").arg(columnLetters[col]).arg(excelRow);
                        worksheet->querySubObject("Range(const QString&)", cell)->setProperty("Value", QString(amountDigits[col]));
                    }
                }
            }

            if (fileIndex == fileCount - 1) {
                // 在最后一个文件中添加总计信息
                double totalSum = 0.0;
                for (const QVariantList &row : qlInvoiceData) {
                    if (row.size() < 5) continue;
                    totalSum += row[4].toDouble() - row[5].toDouble();
                }
                // 在 D22 写入 "合计（小写）"
                worksheet->querySubObject("Range(const QString&)", "D22")->setProperty("Value", "合计（小写）");
                // 在 M22-R22 填充 totalSum 的千位到小数点后两位
                QString totalSumStr = QString::number(totalSum, 'f', 2).remove('.');
                QList<QChar> sumDigits;
                for (const QChar &ch : totalSumStr) {
                    sumDigits.append(ch);
                }
                QStringList totalColumns = {"M22", "N22", "O22", "P22", "Q22", "R22"};
                // 如果 sumDigits 长度小于 totalColumns 长度，则在前面补零
                while (sumDigits.size() < totalColumns.size()) {
                    sumDigits.prepend('0'); // 在开头补零
                }
                int n = totalColumns.size();
                for (int i = 0; i < totalColumns.size(); ++i) {
                    QChar digit = sumDigits.at(i);
                    QString valueToFill;
                    if (i == 0 && digit == '0') {
                        // 千位且为 0，填空白
                        valueToFill = "";
                        n = n - 1;
                    } else if (i == 1 && digit == '0' && n == totalColumns.size() - 1) {
                        valueToFill = "";
                        n = n - 1;
                    } else if (i == 2 && digit == '0' && n == totalColumns.size() - 2) {
                        valueToFill = "";
                        n = n - 1;
                    } else if (i == totalColumns.size() - 3 && digit == '0') {
                        // 个位且为 0，填入 0
                        valueToFill = "0";
                    } else {
                        // 其他位置按实际数字填入
                        valueToFill = QString(digit);
                    }
                    worksheet->querySubObject("Range(const QString&)", totalColumns[i])->setProperty("Value", valueToFill);
                }
                // 在 D23 写入 "合计（大写）：  壹仟叁佰玖拾捌元零角捌分"
                QString chineseNumber = convertToChineseCurrency(totalSum); // 自定义函数
                worksheet->querySubObject("Range(const QString&)", "D23")->setProperty("Value", "合计（大写）：  " + chineseNumber);

                // 在 E24, H24, L24 添加签字栏信息
                worksheet->querySubObject("Range(const QString&)", "E24")->setProperty("Value", "会计审核：");
                worksheet->querySubObject("Range(const QString&)", "H24")->setProperty("Value", "负责人：");
                worksheet->querySubObject("Range(const QString&)", "L24")->setProperty("Value", "验收人：");
                worksheet->querySubObject("Range(const QString&)", "Q24")->setProperty("Value", "经办人：");
            }



            // 保存文件
            QString defaultSaveDir = QStandardPaths::writableLocation(QStandardPaths::DesktopLocation);
            QString saveFileName = QFileDialog::getSaveFileName(this, "保存文件", defaultSaveDir + QString("/验收单%1.xlsx").arg(fileIndex + 1), "Excel 文件 (*.xlsx)");
            if (!saveFileName.isEmpty()) {
                workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(saveFileName));
            }

            workbook->dynamicCall("Close(Boolean)", false); // 关闭工作簿

            // 删除临时文件
            QFile::remove(tempFilePath);
        }

        excel.dynamicCall("Quit()"); // 退出Excel
        tempDir.removeRecursively(); // 删除临时文件夹

        QMessageBox::information(this, "完成", "文件已生成并保存！");
    } catch (const std::exception &e) {
        QMessageBox::critical(this, "错误", e.what());
    }
}

QString MainWindow::convertToChineseCurrency(double amount) {
    static const QStringList numbers = {"零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"};
    static const QStringList units = {"", "拾  ", "佰  ", "仟  "};
    static const QStringList bigUnits = {"元  ", "万  ", "亿  "};

    QString result;
    QString integerPart = QString::number(static_cast<int>(amount));
    QString fractionalPart = QString::number(amount, 'f', 2).split(".")[1];

    // 整数部分处理
    int groupCount = 0;
    while (!integerPart.isEmpty()) {
        QString group = integerPart.right(4);
        integerPart.chop(4);
        QString groupResult;
        for (int i = 0; i < group.size(); ++i) {
            int digit = group[group.size() - 1 - i].digitValue();
            groupResult.prepend(numbers[digit] + units[i]);
        }
        if (!groupResult.isEmpty()) {
            groupResult.append(bigUnits[groupCount]);
        }
        result.prepend(groupResult);
        ++groupCount;
    }

    // 小数部分处理
    if (fractionalPart == "00") {
        result.append("整");
    } else {
        if (fractionalPart[0] != '0') {
            result.append(numbers[fractionalPart[0].digitValue()] + "角  ");
        }
        if (fractionalPart[1] != '0') {
            result.append(numbers[fractionalPart[1].digitValue()] + "分  ");
        }
    }

    return result;
}


MainWindow::~MainWindow() {}
