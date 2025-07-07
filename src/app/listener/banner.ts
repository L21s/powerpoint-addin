import {
    addBannerButton,
    removeBannerButton,
    bannerTextInput,
    bannerTextColorInput,
    bannerBackgroundColorInput,
    bannerPositionSelect,
} from "../taskpane";
import {addBanner, checkBannerExists, removeBanner} from "../actions/banner";

export function initializeBannerListener() {
    addBannerButton.addEventListener("click", async () => {
        const text = bannerTextInput.value;
        const textColor = bannerTextColorInput.value;
        const backgroundColor = bannerBackgroundColorInput.value;
        const position = bannerPositionSelect.value as "Top" | "Left" | "Right";

        await addBanner({ text, textColor, backgroundColor, position });
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
