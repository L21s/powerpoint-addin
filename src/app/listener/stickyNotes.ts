import {stickyNotes} from "../taskpane";
import {insertSticker} from "../actions/stickyNotes";

export function initializeStickyNotesListener() {
    stickyNotes.forEach((button) => {
        const color = button.getAttribute("data-color");
        (button as HTMLElement).onclick = () => insertSticker(color);
    });
}