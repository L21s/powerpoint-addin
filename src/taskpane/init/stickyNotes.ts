import {stickyNotes} from "../taskpane";
import {insertSticker} from "../listeners/stickyNotes";

export function initializeStickyNotes() {
    stickyNotes.forEach((button) => {
        const color = button.getAttribute("data-color");
        (button as HTMLElement).onclick = () => insertSticker(color);
    });
}