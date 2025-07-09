import {resetSearchInputAndDrawer} from "./searchDrawer";
import {employeesPreview} from "../taskpane";
import {Employee} from "../shared/types";
import {fetchEmployeeImage, fetchEmployeeNames} from "../services/employeeApiService";
import {ShapeType} from "../shared/consts";

let employees: Employee[] = [];

export async function getAllEmployeeNames() {
  const employeeNames = await fetchEmployeeNames();

  employees = employeeNames.map((employee) => ({
    id: employee,
    name:
        employee.split("-")[1].charAt(0).toUpperCase() +
        employee.split("-")[1].slice(1) +
        " " +
        employee.split("-")[0].charAt(0).toUpperCase() +
        employee.split("-")[0].slice(1),
  }));
}

export async function fetchEmployeesAddToPreview(searchTerm: string){
  let result = searchTerm ? filterEmployeeNames(searchTerm) : employees;
  result.sort((a, b) => a.name.localeCompare(b.name));
  addToEmployeesPreview(result);
}

function addToEmployeesPreview(names: Employee[]) {
  document.querySelectorAll("sl-skeleton").forEach((skeletonItem) => skeletonItem.remove());

  names.forEach((name) => {
    const menuItem = document.createElement("sl-menu-item") as HTMLButtonElement
    menuItem.id = name.id;
    menuItem.innerText = name.name;

    employeesPreview.appendChild(menuItem);
    menuItem.onclick = (e) => insertEmployeeImage(e, name.id);
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
  resetSearchInputAndDrawer();
}

function filterEmployeeNames(searchTerm: string) {
  return employees.filter((employee) => employee.name.toLowerCase().includes(searchTerm.toLowerCase()));
}
