import fetch from "node-fetch";
import tar from "tar-fs";
import zlib from "zlib";
import fs from "fs/promises";
import path from "path";
import yaml from "yaml";
import Jimp from "jimp";
import * as Rimraf from "rimraf";
import {logErrorMessage} from "office-addin-cli";

// Define the border size for the employee images
const borderSize = 10;

interface EmployeeDataSource {
    github: EmployeeGithubDataSource,
}

interface EmployeeGithubDataSource {
    accessToken: string;
    organization: string;
    repository: string;
    branch: string;
}

interface EmployeeData {
    id: string,
    name: string;
    imageData: Buffer;
}

const employeeDataSourceFile = `${__dirname}/.employee-data-source.json`;

/**
 * Load and validate the data source definition from .employee-data-source.json
 */
async function loadDataSource(): Promise<EmployeeDataSource> {
    function checkKeyExists<T>(name: string, val: T | null | undefined): T {
        if (typeof(val) === "undefined" || val === null) {
            throw new Error(
                `Missing "${name}" in ${employeeDataSourceFile}`
            );
        }

        return val;
    }

    function checkKeyIsString<T>(name: string, val: T | null | undefined): T {
        const valNonNull = checkKeyExists<T>(name, val);

        const valType = typeof(valNonNull);
        if (valType !== "string") {
            throw new Error(`"${name}" is not a string in ${employeeDataSourceFile}`);
        }

        return valNonNull;
    }

    function checkKeyIsObject<T>(name: string, val: T | null | undefined): T {
        const valNonNull = checkKeyExists<T>(name, val);

        const valType = typeof(valNonNull);
        if (valType !== "object") {
            throw new Error(`"${name}" is not an object in ${employeeDataSourceFile}`);
        }

        return valNonNull;
    }

    const employeeDataContent = await fs.readFile(
        `${__dirname}/.employee-data-source.json`,
        { encoding: "utf-8" }
    ).catch((error) => {
        throw new Error(
            "Failed to read .employee-data-source.json\n" +
            `Please create this file at "${employeeDataSourceFile}".\n` +
            "See .employee-data-source.json.template as an example.\n" +
            "\n" +
            `(Original error was: ${error})`
        );
    });

    let employeeData;
    try {
        employeeData = JSON.parse(employeeDataContent);
    } catch (e) {
        throw new Error(`${employeeDataSourceFile} is not valid JSON!`);
    }

    const githubKey = checkKeyIsObject("github", employeeData["github"]);
    const accessToken = checkKeyIsString("github->accessToken", githubKey["accessToken"]);
    const organization = checkKeyIsString("github->organization", githubKey["organization"]);
    const repository = checkKeyIsString("github->repository", githubKey["repository"]);
    const branch = checkKeyIsString("github->branch", githubKey["branch"]);

    return { github: { accessToken, organization, repository, branch } };
}

/**
 * Downloads the data from the github repository.
 *
 * @param dataSource the data source to download from
 * @param outputPath the path to extract the repository to
 */
async function downloadAndExtractGitHubRepo(dataSource: EmployeeDataSource, outputPath: string) {
    const { accessToken, organization, repository, branch } = dataSource.github;

    const branchDownloadUrl = `https://api.github.com/repos/${organization}/${repository}/tarball/${branch}`;

    const response = await fetch(branchDownloadUrl, {
        headers: {
            "Authorization": `token ${accessToken}`,
            "Accept": "application/vnd.github.v3.raw"
        }
    });
    if (!response.ok) {
        throw new Error(
            `Failed to fetch ${branchDownloadUrl}: ${response.statusText} (${response.status}), please check if ${employeeDataSourceFile} is correct`
        );
    }

    return new Promise<void>((resolve, reject) => {
        response.body
            .pipe(zlib.createGunzip())
            .pipe(tar.extract(outputPath, {
                // Skip all files not in a directory
                ignore: (name) => name.indexOf("/") === -1,
                // Strip the repository prefix directory
                map: (header) => {
                    const fileName = header.name;
                    const indexOfSlash = fileName.indexOf("/");
                    header.name = fileName.substring(indexOfSlash + 1);
                    return header;
                }
            }))
            .on("finish", () => {
                console.log(`Repository extracted to: ${outputPath}`);
                resolve();
            })
            .on("error", reject);
    });
}

/**
 * Loads the employee data from the given path.
 *
 * @param repoPath the path to load the data from (expected to be the root of the repository)
 */
async function loadEmployeeData(repoPath: string): Promise<EmployeeData[]> {
    const output = [];

    const directoryEntries = await fs.readdir(repoPath);
    for (const entry of directoryEntries) {
        // Check if its:
        // 1. a directory
        // 2. contains a cv-data.yml
        const employeeDir = path.join(repoPath, entry);
        const stat = await fs.stat(employeeDir);
        if (!stat.isDirectory()) {
            // Not a directory, skip
            continue;
        }

        // Find the cv-data.yml to get the image path
        const employeeDataFile = path.join(employeeDir, "cv-data.yml");
        const employeeDataYaml = await fs.readFile(
            employeeDataFile,
            { encoding: "utf-8" },
        ).catch((error) => {
            if (error?.code !== "ENOENT") {
                // We only expect the file to not exist, other errors are unexpected
                throw error;
            }

            return null;
        });

        if (employeeDataYaml === null) {
            // Not an employee data directory
            continue;
        }

        let employeeRawData;
        try {
            employeeRawData = yaml.parse(employeeDataYaml);
        } catch (e) {
            throw new Error(`${employeeDataFile} is not valid yaml: ${e}`);
        }

        const employeePictureRelPath = employeeRawData["cvPicture"];
        if (typeof(employeePictureRelPath) !== "string") {
            console.warn(`cvPicture in ${employeeDataFile} is not a valid string, skipping this person!`);
            continue;
        }

        const name = employeeRawData["person"]?.["general"]?.["name"] ?? entry;

        let imageData;
        try {
            imageData = await fs.readFile(path.join(employeeDir, employeePictureRelPath));
        } catch (e) {
            throw new Error(`Failed to read employee picture: ${e}`);
        }

        output.push({
            id: entry,
            name,
            imageData,
        })
    }

    return output;
}

/**
 * Processes the employee data and makes it suitable for usage in PowerPoint.
 * 
 * @param data the data to process (mutated in place)
 */
async function processEmployeeData(data: EmployeeData[]) {
    for (let i = 0; i < data.length; i++) {
        const employee = data[i];

        // Load the image
        const image = await Jimp.read(employee.imageData);

        // Make a circle out of the image
        const width = image.getWidth();
        const height = image.getHeight();

        let wantedSize = Math.min(width, height);
        if (width != height) {
            console.warn(`Image is not square for ${employee.id}`);
            image.crop(0, 0, wantedSize, wantedSize);
        }

        // Resize the image to 512x512 (if it isn't already)
        if (wantedSize != 512) {
            const warningMessage = wantedSize < 512
                ? `Image is smaller than 512x512 for ${employee.id} resizing to 512x512, this may result in a blurry image!`
                : `Image is larger than 512x512 for ${employee.id} resizing to 512x512!`;
            console.warn(warningMessage);
            wantedSize = 512;
            image.resize(512, 512);
        }

        const imageCenter = wantedSize / 2;

        for (let y = 0; y < wantedSize; y++) {
            for (let x = 0; x < wantedSize; x++) {
                const xDist = x - imageCenter;
                const yDist = y - imageCenter;

                const distance = Math.sqrt((xDist * xDist) + (yDist * yDist));
                const crop = distance > imageCenter;
                const makeBorder = distance + borderSize > imageCenter;

                if (crop) {
                    image.setPixelColor(Jimp.rgbaToInt(0, 0, 0, 0), x, y);
                } else if (makeBorder) {
                    image.setPixelColor(Jimp.rgbaToInt(82, 55, 252, 255), x, y);
                }
            }
        }

        // Convert to PNG (if it wasn't already)
        employee.imageData = await image.getBufferAsync(Jimp.MIME_PNG);

        console.log(`[${i + 1}/${data.length}] Processed image for ${employee.id}`);
    }
}

async function generateOutput(generationDir: string, data: EmployeeData[]) {
    let code = "const EMPLOYEE_DATA = {\n";


    for (let i = 0; i < data.length; i++) {
        const employee = data[i];

        // Write the output image
        const employeeId = employee.id;
        const imageName = `${employeeId}.png`;
        const imagePath = path.join(generationDir, `${employeeId}.png`);

        await fs.writeFile(imagePath, employee.imageData);

        console.log(`[${i + 1}/${data.length}] Written image for ${employeeId}`);

        // Append the code to the object
        code += `    "${employeeId}": { picture: import("./${imageName}"), name: "${employee.name}", },\n`;
    }
    
    code += "}\n";
    code += "\n";
    code += "export default EMPLOYEE_DATA;\n";

    // Write the generated file
    const dataFile = path.join(generationDir, "employee-data.ts");
    await fs.writeFile(dataFile, code);
}

(async () => {
    const dataSource = await loadDataSource();

    const repositoryPath = await fs.mkdtemp("powerpoint-addin-");
    await Rimraf.rimraf(repositoryPath);
    await fs.mkdir(repositoryPath, { recursive: true });

    console.info("Downloading and extracting employee data...");
    await downloadAndExtractGitHubRepo(dataSource, repositoryPath);
    console.info("... Download finished!");

    console.info("Loading employee data...");
    const data = await loadEmployeeData(repositoryPath);
    console.info( `... Loaded data of ${data.length} employees!`);

    await Rimraf.rimraf(repositoryPath);

    console.info("Processing employee data...");
    await processEmployeeData(data);
    console.info("... Done!");

    const generationDir = path.join(__dirname, "src", "gen");
    await Rimraf.rimraf(generationDir);
    await fs.mkdir(generationDir, { recursive: true });

    console.info("Writing data...");
    await generateOutput(generationDir, data);
    console.info("... Done!");
})();
