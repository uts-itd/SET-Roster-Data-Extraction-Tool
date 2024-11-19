# SET-Roster-Data-Extraction-Tool

The Student Enquiry Team uses data from their rosters which is used in statistics that measure staff performance.
This tool extracts the data from the roster and puts it into a table format which makes it easier to generate statistical information.

This tool is an Excel Add-in written using the Office JavaScript API.

## Development Notes

### Localhost certificates

`localhost` certificates are required for development and testing the add-in on your local machine. These certificates expire after 30 days, after which you will need to reinstall new certificates.

Use the following command to troubleshoot your certificate:

``` bash
npx office-addin-dev-certs verify
```

Use the following command to uninstall the certificate:

``` bash
npx office-addin-dev-certs uninstall --machine
```

Use the following command to install a new certificate:

``` bash
npx office-addin-dev-certs install --machine
```

### Testing the tool

To test the tool, you will need a copy of the SET Roster stored on OneDrive or SharePoint.

Start the test server:

``` bash
npm run dev-server
```

Open the file:

``` bash
npm run start: web -- --document "documentUrl"
```

### Hosting add-in files

This repo is set to public and the files are hosted on GitHub Pages using the `deploy` branch as the source.

### Redeploying a code change

1. Run `git checkout deploy && git merge main --no-edit` to change branch and sync your changes from main  so you are read to redeploy.

2. Now let's bundle our application into `dist` with your build command. For now, that's `npm run build`.

3. Run the following in order:

``` bash
git add dist -f && git commit -m "Deployment commit"
git subtree push --prefix dist origin deploy
git checkout main
```
