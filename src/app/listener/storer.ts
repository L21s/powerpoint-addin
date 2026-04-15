import {
    addStorerButton,
    removeStorerButton,
    storerTextInput,
    storerBackgroundColorInput,
} from "../taskpane";
import {
    addStorer,
    DEFAULT_STORER_BACKGROUND,
    DEFAULT_STORER_TEXT,
    getSavedStorerOptions,
    removeStorer,
    startStorerSync,
} from "../actions/storer";
import {StorerOptions} from "../shared/types";

export function initializeStorerListener() {
    addStorerButton.addEventListener("click", async () => {
        const options: StorerOptions = {
            text: storerTextInput.value || DEFAULT_STORER_TEXT,
            backgroundColor: storerBackgroundColorInput.value || DEFAULT_STORER_BACKGROUND,
        };

        await addStorer(options);
        toggleStorerButtons(true);
    });

    removeStorerButton.addEventListener("click", async () => {
        await removeStorer();
        toggleStorerButtons(false);
    });

    const saved = getSavedStorerOptions();
    if (saved) {
        storerTextInput.value = saved.text;
        storerBackgroundColorInput.value = saved.backgroundColor;
        toggleStorerButtons(true);
        startStorerSync();
    } else {
        toggleStorerButtons(false);
    }
}

function toggleStorerButtons(storerExists: boolean) {
    addStorerButton.style.display = storerExists ? "none" : "inline-block";
    removeStorerButton.style.display = storerExists ? "inline-block" : "none";
}
