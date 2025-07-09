import {
    colButtons,
    createColumnsElement,
    createRowsElement,
    deleteColumnsElement,
    deleteRowsElement,
    rowButtons,
} from "../taskpane";
import {createColumns, createRows, deleteColumns, deleteRows} from "../actions/rowsColumns";

export function initializeRowsColumnsListener() {
    createRowsElement.onclick = () => createRows(+(<HTMLInputElement>document.getElementById("number-of-rows")).value);
    deleteRowsElement.onclick = () => deleteRows();
    rowButtons.forEach((button) => {
        (button as HTMLElement).onclick = () => createRows(Number(button.getAttribute("data-value")));
    });

    createColumnsElement.onclick = () =>
        createColumns(+(<HTMLInputElement>document.getElementById("number-of-columns")).value);
    deleteColumnsElement.onclick = () => deleteColumns();
    colButtons.forEach((button) => {
        (button as HTMLElement).onclick = () => createColumns(Number(button.getAttribute("data-value")));
    });
}