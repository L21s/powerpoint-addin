import { EmployeeName, ShapeType } from "./types";
import { fetchEmployeeImage, fetchEmployeeNames } from "./employeeApiService";
import { getSelectedShapeWith } from "./powerPointUtil";

let allEmployeeNames: EmployeeName[] = [];

export function addToTeamPreview(names: EmployeeName[]) {
  const teamPreviewElement = document.getElementById("team");
  document.querySelectorAll("sl-skeleton").forEach((skeleton) => skeleton.remove());

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

    await context.sync();
  });

  button["loading"] = false;
}

async function getAllEmployeeNames(): Promise<Array<EmployeeName>> {
  const employeeList = await fetchEmployeeNames();
  return employeeList.map((employee) => ({
    id: employee,
    name:
      employee.split("-")[1].charAt(0).toUpperCase() +
      employee.split("-")[1].slice(1) +
      " " +
      employee.split("-")[0].charAt(0).toUpperCase() +
      employee.split("-")[0].slice(1),
  }));
}

export async function filterEmployeeNames(searchTerm: string) {
  const allNames = await getAllEmployeeNames();
  const filteredNames = allNames.filter((name) => name.id.includes(searchTerm.toLowerCase()));
  return searchTerm ? filteredNames : allNames;
}
