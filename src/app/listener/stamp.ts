import {
    addStampButton,
    removeStampButton,
    stampTextInput,
    stampBackgroundColorInput,
    stampPositionSelect,
} from "../taskpane";
import {
    addStamp,
    DEFAULT_STAMP_BACKGROUND,
    DEFAULT_STAMP_POSITION,
    DEFAULT_STAMP_TEXT,
    getSavedStampOptions,
    removeStamp,
    startStampSync,
} from "../actions/stamp";
import {StampPosition} from "../shared/enums";
import {StampOptions} from "../shared/types";

export function initializeStampListener() {
    addStampButton.addEventListener("click", async () => {
        const options: StampOptions = {
            text: stampTextInput.value || DEFAULT_STAMP_TEXT,
            backgroundColor: stampBackgroundColorInput.value || DEFAULT_STAMP_BACKGROUND,
            position: (stampPositionSelect.value as StampPosition) || DEFAULT_STAMP_POSITION,
        };

        await addStamp(options);
        showRemoveStampControls();
    });

    removeStampButton.addEventListener("click", async () => {
        await removeStamp();
        showAddStampControls();
    });

    const saved = getSavedStampOptions();
    if (saved) {
        stampTextInput.value = saved.text;
        stampBackgroundColorInput.value = saved.backgroundColor;
        stampPositionSelect.value = saved.position;
        showRemoveStampControls();
        startStampSync();
    } else {
        showAddStampControls();
    }
}

function showAddStampControls() {
    addStampButton.style.display = "inline-block";
    removeStampButton.style.display = "none";
}

function showRemoveStampControls() {
    addStampButton.style.display = "none";
    removeStampButton.style.display = "inline-block";
}
