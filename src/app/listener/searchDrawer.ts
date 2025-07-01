import {activeDrawer, searchInput} from "../taskpane";
import {closeDrawer, handleDrawerChange, handleSearchInput} from "../actions/searchDrawer";
import {getMsalApp, loginWithDialog} from "../../security/authClient";

export function initializeSearchDrawerListener(){
    initializeDrawer()
    initializeSearchInput()
}

function initializeDrawer() {
    activeDrawer.addEventListener("sl-change", async (e) => {
        let activeAccount = getMsalApp().getActiveAccount();

        if (!activeAccount) {
            activeAccount = await loginWithDialog();
        }

        if(activeAccount) {
            await handleDrawerChange(e);
        }
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