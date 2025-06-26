import { loginWithDialog } from "../security/authClient";
import {registerDrawer, registerSearchInput} from "./ui/searchDrawer";
import {registerLogoDropdownOptions} from "./ui/logoDropdown";
import {initRowsAndColumnsButtons} from "./ui/rowsColumns";
import {initStickerButtons} from "./ui/stickers";
import {registerImageBackgroundEditor} from "./ui/imageBackgroundEditor";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initializeUI();
  }
});

function initializeUI() {
  initStickerButtons()
  initRowsAndColumnsButtons()
  registerDrawer()
  registerSearchInput()
  registerImageBackgroundEditor()
  registerLogoDropdownOptions()
}
