import { EmployeeName } from "./types";
import { fetchEmployeeImage, fetchEmployeeNames } from "./employeeApiService";

let allEmployeeNames: EmployeeName[] = [];

export function addToTeamPreview(names: EmployeeName[]) {
  const teamPreviewElement = document.getElementById("team");

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

  Office.context.document.setSelectedDataAsync(
    // setSelectedDataAsync does not accept "data:image/png;base64," part of the base64 string -> remove it with split
    await fetchEmployeeImage(name),
    { coercionType: Office.CoercionType.Image },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Action failed. Error: " + asyncResult.error.message);
      }
    }
  );

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
