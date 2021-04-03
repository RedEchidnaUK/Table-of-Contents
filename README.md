# Table of Contents Web Part

## Summary
This SharePoint Framework web part displays a table of contents for the current page and is based on [Dzmitry Rogozhny's](https://github.com/dmitryrogozhny) excellent [Table of Contents](https://github.com/dmitryrogozhny/sharepoint-lab/blob/master/table-of-contents/). I haven't forked the code as it was originally for an internal project and I didn't want to fork the entire repositry, just this sub folder. There's probably some way in GIT to fork just a sub folder, but I don't know how.


![web part preview](./assets/table-of-contents-display.jpg)

### Web part properties

![web part properties](./assets/table-of-contents-properties.jpg)

The web part provides the following properties:
- `Show Headings 1`&thinsp;&mdash;&thinsp;whether to show H1 headings.
- `Show Headings 2`&thinsp;&mdash;&thinsp;whether to show H2 headings.
- `Show Headings 3`&thinsp;&mdash;&thinsp;whether to show H3 headings.
- `Show the previous page link`&thinsp;&mdash;&thinsp;whether to show a link to the previous page. You can use this to just have a 'Return to previous page' link if you set the title to a space character so it doesn't appear
- `Enable Sticky Mode`&thinsp;&mdash;&thinsp;Makes the table of contents 'stick' to where it was on the screen when a long page is scrolled*.
- `Hide in mobile view`&thinsp;&mdash;&thinsp;whether to hide the web part on small screens.
- `Title`&thinsp;&mdash;&thinsp;allows to specify the title for the web part. You can edit the title directly in the web part's body.

*This mode will not work correctly on the local workbench, only the live site. It should also be used with caution as it works by manipulating Microsoft's styles on the containing element, so it may stop working if Microsoft change their code, you have been warned!

## Minimal Path to Awesome

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
