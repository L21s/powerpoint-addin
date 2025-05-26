import { EmployeeName, ShapeType } from "./types";
import { fetchEmployeeImage, fetchEmployeeNames } from "./employeeApiService";

export let allCurrentNames: EmployeeName[] = [];

export function addToTeamPreview(names: EmployeeName[]) {
  const teamPreviewElement = document.getElementById("names");
  document.querySelectorAll("sl-skeleton").forEach((skeletonItem) => skeletonItem.remove());

  names.forEach((name) => {
    const menuItemElement = document.createElement("sl-menu-item") as HTMLButtonElement;
    menuItemElement.id = name.id;
    menuItemElement.innerText = name.name;

    teamPreviewElement.appendChild(menuItemElement);
    menuItemElement.onclick = (e) => insertEmployeeImage(e, name.id);
  });
}

async function insertEmployeeImage(e: MouseEvent, name: string) {
  const button = e.target as HTMLButtonElement;
  button["loading"] = true;

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const background = slide.shapes.addGeometricShape(ShapeType["Ellipse"]);

    background.width = 100;
    background.height = 100;
    background.fill.setImage(await fetchEmployeeImage(name));
    background.lineFormat.weight = 2;
    background.lineFormat.color = "#5237fc";

    await context.sync();
  });

  button["loading"] = false;
}

export async function getAllEmployeeNames() {
  const employeeList = await fetchEmployeeNames();
  allCurrentNames = employeeList.map((employee) => ({
    id: employee,
    name:
      employee.split("-")[1].charAt(0).toUpperCase() +
      employee.split("-")[1].slice(1) +
      " " +
      employee.split("-")[0].charAt(0).toUpperCase() +
      employee.split("-")[0].slice(1),
  }));
}

export function filterEmployeeNames(searchTerm: string) {
  return allCurrentNames.filter((name) => name.id.includes(searchTerm.toLowerCase()));
}
