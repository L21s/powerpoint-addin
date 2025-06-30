import {activeDrawer, searchInput} from "../taskpane";
import {closeDrawer, handleDrawerChange, handleSearchInput} from "../actions/searchDrawer";

export function initializeSearchDrawerListener(){
    initializeDrawer()
    initializeSearchInput()
}

function initializeDrawer() {
    activeDrawer.addEventListener("sl-change", async (e) => {
        await handleDrawerChange(e)
    });

    document.getElementById("close-drawer").onclick = () => {
        closeDrawer()
    };
}

function initializeSearchInput() {
    searchInput.addEventListener("sl-input", () => {
        handleSearchInput()
    });
}