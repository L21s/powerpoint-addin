import {
    colButtons,
    createColumnsElement,
    createRowsElement,
    deleteColumnsElement,
    deleteRowsElement,
    rowButtons,
} from "../taskpane";
import {columnLineName, createColumns, createRows, deleteShapesByName, rowLineName} from "../listeners/rowsColumns";

export function initializeRowsColumns() {
    createRowsElement.onclick = () => createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
    deleteRowsElement.onclick = () => deleteShapesByName(rowLineName);
    rowButtons.forEach((button) => {
        (button as HTMLElement).onclick = () => createRows(Number(button.getAttribute("data-value")));
    });

    createColumnsElement.onclick = () =>
        createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
    deleteColumnsElement.onclick = () => deleteShapesByName(columnLineName);
    colButtons.forEach((button) => {
        (button as HTMLElement).onclick = () => createColumns(Number(button.getAttribute("data-value")));
    });
}