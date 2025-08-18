import {activeDrawer, searchInput} from "../taskpane";
import {closeDrawer, handleDrawerChange, handleSearchInput} from "../actions/searchDrawer";
import {getActiveAccount, loginWithDialog} from "../services/authService";

export function initializeSearchDrawerListener(){
    initializeDrawer()
    initializeSearchInput()
}

function initializeDrawer() {
    activeDrawer.addEventListener("sl-change", async (e) => {
        let activeAccount = getActiveAccount();

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
    searchInput.addEventListener("sl-input", async () => {
        await handleSearchInput()
    });
}