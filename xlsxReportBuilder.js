// const excel = require("excel4node");
// const fs = require("fs");
const XLSX = require("xlsx-js-style");

// You can define styles as json object

// у xlsx-js-style нет метода workbook.createStyle(), поэтому стили будут чуть другие, чем у excel4Node
//Что было изменено в стиле ниже, чтобы это подходило для xlsx-js-style
//1. У fgcolor убраны первые FF и добавлен внутрь объект, так как клетки становились полностью черными
//TODO: нужно поменять все цвета на rgb, подозреваю, что тут везде так
//2. Размер шрифтов называется теперь не size, он был поменян на sz
//3. Формат числа теперь не numberFormat, он был поменян на numFmt
const styles = {
    headerGrey: {
        fill: {
            type: "pattern",
            patternType: "solid",
            fgColor: {rgb: "DEE6EF"},
        },
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
        font: {
            color: "FF000000",
            name: "Arial",
            sz: 10,
            bold: true,
            underline: false,
        },
        alignment: {
            vertical: "top",
            horizontal: "center",
        },
        wrapText: true,
    },
    cellNum: {
        numFmt: "#,##0",
        font: {name: "Arial", sz: 10},
        alignment: {
            vertical: "top",
            horizontal: "right",
        },
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellPercent: {
        numFmt: "0.0%",
        font: {name: "Arial", sz: 10},
        alignment: {
            vertical: "top",
            horizontal: "right",
        },
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellCenter: {
        alignment: {
            vertical: "top",
            horizontal: "center",
        },
        numFmt: "0",
        font: {name: "Arial", sz: 10},
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellQuantity: {
        alignment: {
            vertical: "top",
            horizontal: "center",
        },
        numFmt: "0.0##",
        font: {name: "Arial", sz: 10},
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellDate: {
        alignment: {
            vertical: "top",
            horizontal: "center",
        },
        numFmt: "dd.mm.yy",
        font: {name: "Arial", sz: 10},
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellDateTime: {
        alignment: {
            vertical: "top",
            horizontal: "center",
        },
        numFmt: "yyyy-mm-dd hh:mm",
        font: {name: "Arial", sz: 10},
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
    cellDefault: {
        alignment: {
            vertical: "top",
        },
        font: {name: "Arial", sz: 10},
        border: {
            top: {style: "thin", color: "#404040"},
            bottom: {style: "thin", color: "#404040"},
            left: {style: "thin", color: "#404040"},
            right: {style: "thin", color: "#404040"},
        },
    },
}

function getXLSX(data) {
    const report = buildExport(data.sheets);
    if (report) return report; //toBuffer метода нет поэтому прокидываю toString
}

function getTypeByValue(value) { //ОСТОРОЖНО! Данная функция работает только для xlsx-js-style
    if (value == null) value = undefined;
    switch (typeof value) {
        case "number":
            return "n";
        case "string":
            return "s";
        case "undefined":
            return "s";
        // case "undefined":
        // 	cell.string("");
        // 	break;
        default:
            if (typeof value.getMonth === "function") {
                return 'd';
            }
            return 's'
    }
}

function buildExport(sheets) {
    const workBook = XLSX.utils.book_new();
    let stylebook = {};

    //Для xlsx-js-style
    Object.keys(styles).forEach(stylename => {
        stylebook[stylename] = styles[stylename];
    });

    sheets.forEach(sheet => {
        if (!sheet.specification) return;
        let heading = sheet.heading || [];

        let subRowOfHeadings = []; //массив ячеек заголовка одного листа
        heading.forEach((r) => {
            if (r instanceof Array) {
                r.forEach((val) => {
                    let m = {};
                    if (val !== null) {
                        m.v = val;
                        m.t = getTypeByValue(val);
                    } else {
                        m.value = val;
                    }
                    subRowOfHeadings.push(m);
                });
            }
        });

        let rowOfColnames = []; // ряд заголовков каждого столбца
        Object.keys(sheet.specification).forEach((colname) => {
            let spec = sheet.specification[colname];  //получаем соотвествующее название
            let name = spec.displayName.toString();
            let cell = {v: name, t: getTypeByValue(name)};
            if (stylebook[spec.headerStyle]) {
                cell.s = stylebook[spec.headerStyle];
            }
            rowOfColnames.push(cell)
        });

        //создаем массив ячеек, куда по одной ячейке будем загонять и создавать строку таблицы
        //в store хранятся данные ячеек
        let store = [];

        // объединить ячейки в xlsx-js-style можно только 1 раз, поэтому будем копить объединения тут
        let merges = [];

        //ниже переменные нужны для успешного смерживания
        let temporaryStore = [];
        let subMerge = [];
        let inRowNoMerges = false;

        let i = 0; //TEST
        let flag = false;

        sheet.data.forEach((row, rowno) => {

            if (flag) {
                return 0;
            }

            if (inRowNoMerges) {
                merges.push(...temporaryStore);
                temporaryStore = [];
            } else {
                inRowNoMerges = true;
                temporaryStore = subMerge.slice(0);
            }
            store.push([]);
            subMerge = [];//обнуляем subMerge перед перебором ряда
            row._row_number = rowno + 1;
            Object.keys(sheet.specification).forEach((colname, colno) => {
                let value = row[colname];
                let spec = sheet.specification[colname];
                let res = {};
                if (spec.styleFunc && typeof spec.styleFunc === "function") {

                }
                if (spec.beforeWrite && typeof spec.beforeWrite === "function") {
                    res = spec.beforeWrite(value, {
                        dataset: sheet.data,
                        row,
                        rowno,
                        colname,
                    });
                    value = res.newvalue;
                }


                if (value === null) value = "";
                let m = {
                    v: value,
                    t: getTypeByValue(value),
                    s: stylebook[spec.cellStyle]
                };
                store[rowno].push(m); //вставляем ячейку в rowno ряд

                if (res.merges) {            // проверка на то, что ячейка дожна быть смержена
                    inRowNoMerges = false;
                    subMerge.push({
                        //+3 захардкожено, потому что есть отступы в первом листе из-за заголовков
                        s: {r: rowno - res.merges.up + 3, c: colno - res.merges.left},
                        e: {r: rowno + 3, c: colno}
                    })
                }

                i++;
                if (i >= 1000) {
                    flag = true;
                    return 0;
                }
            });
        });

        //кастомизация каждого листа отдельно
        let arrOfWidths = []; //ширина столбцов
        let rowOfHeadings = [];
        if (sheet.name === 'Поставщики') {
            arrOfWidths = [
                {width: 4}, //sid
                {width: 20}, //Поставщик
                {width: 7}, //Ликвид
                {width: 12}, //Магазинов
                {width: 11}, //Артикулов 1
                {width: 11}, //Выручка 1
                {width: 7}, //Доля 1
                {width: 15}, //Обновлено
                {width: 10}, //Маг2
                {width: 11}, //Артикулов2
                {width: 11}, //Выручка 2
                {width: 7}, //Доля2
            ];
            rowOfHeadings.push(subRowOfHeadings, [])
        } else if (sheet.name === 'Товары 80% выручки') {
            arrOfWidths = [
                {width: 4}, //id
                {width: 15}, //Поставщик
                {width: 20}, //Штрихкод
                {width: 35}, //Товар
                {width: 10}, //Цена мин
                {width: 10}, //Цена макс
                {width: 20}, //Посл поставщик
                {width: 10}, //Посл цена
                {width: 12}, //Продажи шт.
                {width: 10}, //Закупка
                {width: 11}, //Магазинов
                {width: 35}, //Магазины
                {width: 20}, //Категория
            ];
        }

        let workSheet = XLSX.utils.aoa_to_sheet([...rowOfHeadings, rowOfColnames, ...store]);//создаем лист (worksheet) и запихиваем туда данные
        console.log(merges);
        workSheet["!merges"] = merges; //делаем смерживание

        //Ширина столбцов применяется
        workSheet["!cols"] = arrOfWidths;

        XLSX.utils.book_append_sheet(workBook, workSheet, sheet.name); //передаем книге (workbook) наш лист и называем его sheet.name
        XLSX.writeFile(workBook, "demka.xlsx");
    });
    return workBook;
}
module.exports = {getXLSX, styles};