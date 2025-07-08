import {
    addBannerButton,
    removeBannerButton,
    bannerTextInput,
    bannerTextColorInput,
    bannerBackgroundColorInput,
    bannerPositionSelect,
} from "../taskpane";
import {addBanner, checkBannerExists, removeBanner} from "../actions/banner";
import {BannerPosition} from "../shared/enums";
import {BannerOptions} from "../shared/types";

export function initializeBannerListener() {
    addBannerButton.addEventListener("click", async () => {
        const bannerOptions: BannerOptions = {
            text: bannerTextInput.value,
            textColor: bannerTextColorInput.value,
            backgroundColor: bannerBackgroundColorInput.value,
            position: bannerPositionSelect.value as BannerPosition
        }

        await addBanner(bannerOptions);
        toggleBannerButtons(true);
    });

    removeBannerButton.addEventListener("click", async () => {
        await removeBanner();
        toggleBannerButtons(false);
    });

    checkBannerExists().then(toggleBannerButtons);
}

function toggleBannerButtons(bannerExists: boolean) {
    addBannerButton.style.display = bannerExists ? "none" : "inline-block";
    removeBannerButton.style.display = bannerExists ? "inline-block" : "none";
}
