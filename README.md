### Prerequisites

1. **Install Node.js**: Ensure Node.js is installed on your system. You can download it from [nodejs.org](https://nodejs.org/).

2. **Install Office Add-in CLI**: Open a terminal and run the following command to install the Office Add-in CLI globally:
    ```bash
    npm install -g yo generator-office
    ```

3. **Install Visual Studio Code**: Download and install [Visual Studio Code](https://code.visualstudio.com/).

4. **Install Office Add-in Debugger Extension**: In Visual Studio Code, go to the Extensions view (Ctrl+Shift+X) and install the "Office Add-in Debugger" extension.

### About this project

This project uses `Excel Custom Functions using a JavaScript-only Runtime`, the script type is `JavaScript` and the project type is `Excel Add-in`.

### Explore the Project Structure

The generated project will have the following structure:

```plaintext
your-project-name/
├── .vscode/
├── node_modules/
├── src/
│   ├── commands/
│   ├── taskpane/
│   ├── assets/
│   ├── functions/
│   └── styles/
├── manifest.xml
├── package.json
└── webpack.config.js
```

### Build and Run the Project

1. **Install Dependencies**:

    ```bash
    npm install
    ```

2. **Start the Local Web Server a local excel file**:

    ```bash
    npm start
    ```

Or

2. **Start the Local Web Server running a web browser excel file**:

    ```bash
    npm run start:web -- --document "https://url of your Excel document"
    ```

3. **Sideload the Add-in**:
    - Open Excel.
    - Go to the `Insert` tab.
    - Click `My Add-ins` > `Manage My Add-ins` > `Upload My Add-in`.
    - Select the `manifest.xml` file from your project directory.

### Develop Your Add-in

1. **Open the Task Pane**: After sideloading, you should see your add-in's task pane. You can now start developing your add-in.

2. **Edit Code**: Open `src/taskpane/taskpane.html` and `src/taskpane/taskpane.js` in Visual Studio Code to customize your add-in's UI and functionality.

### Debug Your Add-in

1. **Set Breakpoints**: Open the files in Visual Studio Code and set breakpoints in your JavaScript code.

2. **Run the Debugger**: Press `F5` to start debugging. This will attach the debugger to Excel, and you can inspect variables, step through code, and troubleshoot issues.

### Publish Your Add-in

1. **Prepare for Production**: Modify your add-in as needed for production. Ensure you have proper icons, descriptions, and a polished UI.

2. **Host the Web App**: Deploy your web app to a hosting service (e.g., Azure, AWS, Heroku).

3. **Update Manifest**: Update the `manifest.xml` file with the production URL of your hosted web app.

4. **Distribute the Add-in**: You can distribute the `manifest.xml` file directly to users or submit your add-in to the Office Store.

### Additional Resources

- [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API](https://docs.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview)
- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js](https://nodejs.org/)

This guide should help you get started with building an Excel add-in using Visual Studio Code and Node.js. If you have any specific questions or need further assistance, feel free to ask!
