import { fetchIconsAndAddToPreview } from "./iconsPreview";
import {fetchEmployeesAddToPreview, getAllEmployeeNames} from "./employeesPreview";
import {activeDrawer, drawer, searchInput, wrapper} from "../taskpane";

let lastSearchQuery = "";

export async function handleDrawerChange(e: Event) {
  const activeDrawerTab = e.target as HTMLInputElement;

  refreshSearchResults(activeDrawerTab.value);

  drawer["open"] = true;
  wrapper.style.overflow = "hidden";
  wrapper.scrollTo({
    top: 0,
    behavior: "smooth",
  });

  const currentSearchQuery = searchInput.value;
  searchInput.setAttribute("placeholder", "search " + activeDrawerTab.value + "...");
  searchInput.focus();
  searchInput.value = lastSearchQuery;
  lastSearchQuery = currentSearchQuery;

  const tabs = document.querySelector("sl-split-panel") as any;
  switch (activeDrawerTab.value) {
    case "icons": {
      tabs.position = 100;
      break;
    }
    case "names": {
      tabs.position = 0;
      await getAllEmployeeNames();
      break;
    }
  }

  await processInputChanges(activeDrawerTab.value);
}

export async function handleSearchInput() {
  refreshSearchResults(activeDrawer.value);
  await processInputChanges(activeDrawer.value);
}

export function closeDrawer() {
  resetSearchInputAndDrawer();
  wrapper.style.overflow = "scroll";
}

export function resetSearchInputAndDrawer() {
  drawer["open"] = false;
  searchInput.value = "";
  activeDrawer.value = "";
}

function refreshSearchResults(activeDrawerTab: string) {
  if (activeDrawerTab) {
    document.getElementById(activeDrawerTab).replaceChildren();

    (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "block";

    for (let i = 0; i < 12; i++) {
      const skeleton = document.createElement("sl-skeleton");
      skeleton.classList.add(activeDrawerTab);
      skeleton.setAttribute("effect", "pulse");
      document.getElementById(activeDrawerTab).appendChild(skeleton);
    }
  }
}

async function processInputChanges(activeDrawerTab: string) {
  const searchResultTitle = document.getElementById(activeDrawerTab + "-search-title");

  try {
    switch (activeDrawerTab) {
      case "icons": {
        await fetchIconsAndAddToPreview(searchInput.value);
        searchResultTitle.innerText = searchInput.value ? 'Search results for "' + searchInput.value + '"' : "Recently used icons";
        if (document.getElementById(activeDrawerTab).children.length === 0) {
          showMessageInDrawer("No recent icons yet");
        }
        break;
      }
      case "names": {
        await fetchEmployeesAddToPreview(searchInput.value);
        searchResultTitle.innerText = searchInput.value ? 'Search results for "' + searchInput.value + '"' : "All employees";
        if (document.getElementById(activeDrawerTab).children.length === 0) {
          showMessageInDrawer("No names fitting this search query");
        }
        break;
      }
    }
  } catch (e) {
    showMessageInDrawer("Could not fetch any " + activeDrawerTab + ": " + e.message);
  }
  (document.querySelector("#search-input > sl-spinner:first-of-type") as HTMLElement).style.display = "none";
}

function showMessageInDrawer(message: string) {
  const iconPreviewElement = document.getElementById(activeDrawer.value);
  const textElement = document.createElement("div");
  textElement.classList.add("information", activeDrawer.value);
  textElement.innerText = message;
  iconPreviewElement.appendChild(textElement);
}