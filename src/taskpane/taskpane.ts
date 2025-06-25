import { loginWithDialog } from "../security/authClient";
import {registerDrawerToggle, registerSearch} from "./ui/searchDrawer";
import {registerIconBackgroundTools} from "./utils/iconUtils";
import {registerLogoImageInsert} from "./ui/logoDropdown";
import {initRowsAndColumnsButtons} from "./ui/rowsColumns";
import {initStickerButtons} from "./ui/stickers";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loginWithDialog();
    initializeUI();
  }
});

function initializeUI() {
  initStickerButtons()
  initRowsAndColumnsButtons()
  registerDrawerToggle()
  registerSearch()
  registerIconBackgroundTools()
  registerLogoImageInsert()
}
