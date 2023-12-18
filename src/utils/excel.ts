import { ParsingOptions, Sheet2JSONOpts, WorkBook, WorkSheet } from "xlsx";

/**
 * Преобразование объекта типа File в XLSX.WorkBook
 * @param file файл с расширением .xls, .xlsx
 * @param opts настройки парсинга файла | см. https://docs.sheetjs.com/#parsing-options
 */
export async function getExcelFromfile(file: File, opts?: ParsingOptions): Promise<WorkBook> {
    const { read } = await import("xlsx");
    return read(await file.arrayBuffer(), opts ?? { type: "array" });
}

/**
 * Получения двумерного массива (только 1ый Worksheet) из файла с расширением .xls, .xlsx
 * @param workBook файл с расширением .xls, .xlsx
 * @param opts настройки парсинга данных из excel | см. https://docs.sheetjs.com/#json
 */
export async function getArrOfExcelFirstWorksheet(workBook: WorkBook, options?: Omit<Sheet2JSONOpts, "header"> & { header: 1 }): Promise<unknown[][]> {
    const { utils } = await import("xlsx");
    const firstWorksheet: WorkSheet = workBook.Sheets[workBook.SheetNames[0]];
    return utils.sheet_to_json(firstWorksheet, options ?? undefined);
}

/**
 * Получение клиентом файла с компьютера пользователя(бинд на любой клик)
 * @param accept перечесление типов файла (строка)
 * @param multiple (заменяет template T) выбор нескольких файлов
 */
export function getFileFromLocal<T extends boolean>(accept: string, multiple: T): Promise<(T extends true ? File[] : File) | null> {
    return new Promise((resolve) => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = accept;
        input.multiple = multiple;

        input.addEventListener("click", () => {
            window.document.body.onfocus = () => {
                setTimeout(() => {
                    window.document.body.onfocus = null;
                    input.parentNode?.removeChild(input);
                    const files = Array.from(input.files ?? []);
                    if (!files.length) return resolve(null);
                    resolve((multiple === true ? files : files[0]) as (T extends true ? File[] : File) | null);
                }, 500);
            };
        });
        input.click();
    });
}

export function readFileAsync(file: File): Promise<{ data: string; filename: string }> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            resolve({
                data: reader.result?.toString().split(",")[1] ?? "",
                filename: file.name.split(".")[0],
            });
        };

        reader.onerror = reject;
    });
}
