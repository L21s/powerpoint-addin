import { loginWithDialog } from "../security/authClient";
import {initializeSearchDrawer} from "./ui/searchDrawer";
import {initializeLogoDropdown} from "./ui/logoDropdown";
import {initializeRowsColumns} from "./ui/rowsColumns";
import {initializeStickyNotes} from "./ui/stickyNotes";
import {initializeImageBackgroundEditor} from "./ui/imageBackgroundEditor";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initializeUI();
  }
});

function initializeUI() {
  initializeStickyNotes()
  initializeRowsColumns()
  initializeSearchDrawer()
  initializeImageBackgroundEditor()
  initializeLogoDropdown()
}
