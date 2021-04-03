# Table of Contents Web Part

## Summary
This SharePoint Framework web part displays a table of contents for the current page and is based on [Dzmitry Rogozhny's](https://github.com/dmitryrogozhny) excellent [Table of Contents](https://github.com/dmitryrogozhny/sharepoint-lab/blob/master/table-of-contents/). I haven't forked the code as it was originally for an internal project and I didn't want to fork the entire repositry, just this sub folder. There's probably some way in GIT to fork just a sub folder, but I don't know how.


![web part preview](./assets/table-of-contents-display.png)

### Web part properties

![web part properties](./assets/table-of-contents-properties.png)

The web part provides the following properties:
- `Hide heading`&thinsp;&mdash;&thinsp;whether to show the main heading.
- `Show Headings 1`&thinsp;&mdash;&thinsp;whether to show H1 headings.
- `Show Headings 2`&thinsp;&mdash;&thinsp;whether to show H2 headings.
- `Show Headings 3`&thinsp;&mdash;&thinsp;whether to show H3 headings.
- `Show the previous page link`&thinsp;&mdash;&thinsp;whether to show a link to the previous page. If used in conjunction with 'Hide heading', you could use this to just have a 'Return to previous page' link. 
- `Enable 'Sticky Mode'`&thinsp;&mdash;&thinsp;Makes the table of contents 'stick' to where it was on the screen when a long page is scrolled*.
- `Hide on small mobile devices`&thinsp;&mdash;&thinsp;whether to hide the web part on small screens.

*This mode will not work correctly on the local workbench, only the live site. It should also be used with caution as it works by manipulating Microsoft's styles on the containing element, so it may stop working if Microsoft change their code, you have been warned!

## Minimal Path to Awesome
### Pre-built package
You can grab the pre-built package ready for deployment from [./release/table-of-contents.sppkg](https://github.com/RedEchidnaUK/Table-of-Contents/blob/master/release/table-of-contents.sppkg).

Here's how to deploy the web part to the site collection's app catalog:
![deploying to site app catalog](./assets/table-of-contents-deploy-to-site-app-catalog.gif)

### Local testing
- Clone this repository
- In the command line run:
  - `npm install`
  - `gulp serve`

### Deploy
- `gulp clean`
- `gulp bundle --ship`
- `gulp package-solution --ship`
- Upload .sppkg file from sharepoint\solution to your tenant App Catalog
  - E.g.: https://<tenant>.sharepoint.com/sites/AppCatalog/AppCatalog
- Add the web part to a site collection, and test it on a page
