import { fetchIconsAndAddToPreview } from "./iconsPreview";
import {fetchEmployeesAddToPreview, getAllEmployeeNames} from "./employeesPreview";

let lastSearchQuery = "";
const debouncedProcessInputChanges = debounce(processInputChanges);
const searchInput = document.getElementById("search-input") as HTMLInputElement;
const drawer = document.getElementById("search-drawer") as HTMLElement;
const activeDrawer = document.getElementById("active-drawer") as HTMLInputElement;
const wrapper = document.getElementById("wrapper") as HTMLElement;

export function registerSearchInput() {
  searchInput?.addEventListener("sl-input", () => {
    refreshSearchResults(activeDrawer.value);
    debouncedProcessInputChanges(activeDrawer.value);
  });
}

export function registerDrawer() {
  activeDrawer.addEventListener("sl-change", async (e) => {
    const activeDrawerTab = (e.target as HTMLInputElement).value;
    refreshSearchResults(activeDrawerTab);

    drawer["open"] = true;
    wrapper.style.overflow = "hidden";
    wrapper.scrollTo({
      top: 0,
      behavior: "smooth",
    });

    const currentSearchQuery = searchInput.value;
    searchInput.setAttribute("placeholder", "search " + activeDrawerTab + "...");
    searchInput.focus();
    searchInput.value = lastSearchQuery;
    lastSearchQuery = currentSearchQuery;

    const tabs = document.querySelector("sl-split-panel") as any;
    switch (activeDrawerTab) {
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
    debouncedProcessInputChanges(activeDrawerTab);
  });

  document.getElementById("close-drawer").onclick = () => {
    drawer["open"] = false;
    wrapper.style.overflow = "scroll";

    searchInput.value = "";
    activeDrawer.value = "";
  };
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
  const searchTerm = searchInput.value;
  const searchResultTitle = document.getElementById(activeDrawerTab + "-search-title");

  try {
    switch (activeDrawerTab) {
      case "icons": {
        await fetchIconsAndAddToPreview(searchTerm);
        searchResultTitle.innerText = searchTerm ? 'Search results for "' + searchTerm + '"' : "Recently used icons";
        if (document.getElementById(activeDrawerTab).children.length === 0) {
          showMessageInDrawer("No recent icons yet");
        }
        break;
      }
      case "names": {
        await fetchEmployeesAddToPreview(searchTerm);
        searchResultTitle.innerText = searchTerm ? 'Search results for "' + searchTerm + '"' : "All employees";
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


function debounce(func: Function) {
  let timer: NodeJS.Timeout;
  return (...args: any[]) => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      func.apply(this, args);
    }, 500);
  };
}

function showMessageInDrawer(message: string) {
  const activeDrawerTab = activeDrawer.value;
  const iconPreviewElement = document.getElementById(activeDrawerTab);
  const textElement = document.createElement("div");
  textElement.classList.add("information", activeDrawerTab);
  textElement.innerText = message;
  iconPreviewElement.appendChild(textElement);
}