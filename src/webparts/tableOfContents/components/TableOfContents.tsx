import * as React from 'react';
import styles from './TableOfContents.module.scss';
import { ITableOfContentsProps } from './ITableOfContentsProps';
import { ITableOfContentsState } from './ITableOfContentsState';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from "TableOfContentsWebPartStrings";

/**
 * Describes a link for a header
 */
interface Link {
  /**
   * The Source html element.
   */
  element: HTMLElement | undefined;
  /**
   * Child nodes for the link.
   */
  childNodes: Link[];
  /**
   * Parent link. Undefined for the root link.
   */
  parent: Link | undefined;
}

export default class TableOfContents extends React.Component<ITableOfContentsProps, ITableOfContentsState> {
  private static timeout = 500;

  private static h2Tag = "h2";
  private static h3Tag = "h3";
  private static h4Tag = "h4";

  /**
   * Create a state for the history count. 
   * This is required to make sure we go back to the correct page when the back to previous page link is clicked.
   */
  constructor(props: ITableOfContentsProps) {
    super(props);
    this.state = {
      historyCount: -1
    };
  }

  /**
   * Gets a nested list of links based on the list of headers specified.
   * @param headers List of HtmlElements for H2, H3, and H4 headers.
   */
  private getLinks(headers: HTMLElement[]): Link[] {
    // create a root link that will be a root for links' tree
    const root: Link = { childNodes: [], parent: undefined, element: undefined };

    let prevLink: Link | null = null;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const link: Link = { childNodes: [], parent: undefined, element: header };

      if (i === 0) {
        // the first header is always added as a child of the root
        link.parent = root;
        root.childNodes.push(link);
      } else {
        const prevHeader = headers[i - 1];

        // compare the current header and the previous one to define where to add new link
        const compare = this.compareHeaders(header.tagName, prevHeader.tagName);

        if (compare === 0) {
          // if headers are on the same level, add header to the same parent
          link.parent = prevLink.parent;
          prevLink.parent.childNodes.push(link);
        } else if (compare < 0) {

          let targetParent = prevLink.parent;
          // if current header bigger than the previous one, go up in the hierarchy to find a place to add link
          // go up in the hierarchy of links until a link with bigger tag is found or until the root link found
          // i.e. for H4 look for H3 or H2, for H3 look for H2, for H2 look for the root.
          while ((targetParent != root) && (this.compareHeaders(header.tagName, targetParent.element.tagName) <= 0)) {
            targetParent = targetParent.parent;
          }

          link.parent = targetParent;
          targetParent.childNodes.push(link);
        } else {
          // if current header is smaller thab the previous one, add link for it as a child of the previous link
          link.parent = prevLink;
          prevLink.childNodes.push(link);
        }
      }

      prevLink = link;
    }

    // return list of links for top-level headers
    return root.childNodes;
  }

  /**
   * Compares two header tags by their weights.
   * The function is used to compare the size of headers (e.g. should H3 go under H2?)
   * @param header1
   * @param header2
   */
  private compareHeaders(header1: string, header2: string): number {
    return this.getHeaderWeight(header1) - this.getHeaderWeight(header2);
  }

  /**
   * Returns a digital weight of a tag. Used for comparing header tags.
   * @param header
   */
  private getHeaderWeight(header: string): number {
    switch (header.toLowerCase()) {
      case (TableOfContents.h2Tag):
        return 2;
      case (TableOfContents.h3Tag):
        return 3;
      case (TableOfContents.h4Tag):
        return 4;
      default:
        throw new Error('Unknown header: ' + header);
    }
  }

  /**
   * Returns html elements in the current page specified by the query selector.
   */
  private getHtmlElements(querySelector: string): HTMLElement[] {
    if (querySelector.length === 0) {
      return [];
    } else {
      const elements = document.querySelectorAll(querySelector);
      const htmlElements: HTMLElement[] = [];

      for (let i = 0; i < elements.length; i++) {
        htmlElements.push(elements[i] as HTMLElement);
      }

      return htmlElements;
    }
  }

  /**
   * Returns a query selector based on the specified props
   * @param props
   */
  private getQuerySelector(props: ITableOfContentsProps) {
    const queryParts = [];
    const classes = ['.cke_editable', '.ck-content', '.Collapsible']

    if (props.showHeading2) {
      for (let i = 0; i < classes.length; i++) {
        queryParts.push(classes[i] + " " + TableOfContents.h2Tag);
      }
    }

    if (props.showHeading3) {
      for (let i = 0; i < classes.length; i++) {
        queryParts.push(classes[i] + " " + TableOfContents.h3Tag);
      }
    }

    if (props.showHeading4) {
      for (let i = 0; i < classes.length; i++) {
        queryParts.push(classes[i] + " " + TableOfContents.h4Tag);
      }
    }

    return queryParts.join(',');
  }

  /**
   * Filters elements with empty text.
   * @param element
   */
  private filterEmpty(element: HTMLElement): boolean {
    return element.innerText.trim() !== '';
  }

  /**
   * Filters elements that are inside <aside> tag and thus not related to a page.
   * @param element
   */
  private filterAside(element: HTMLElement): boolean {
    let inAsideTag = false;

    let parentElement = element.parentElement;

    while (parentElement) {
      if (parentElement.tagName.toLocaleLowerCase() === 'aside') {
        inAsideTag = true;
        break;
      }

      parentElement = parentElement.parentElement;
    }

    return !inAsideTag;
  }

  /**
   * Filters elements that have the data attrribute of 'data-toc-ignore' and thus should be ignored.
   * @param element
   */
  private filterTocIgnore(element: HTMLElement): boolean {
    return !(element.getAttribute("data-toc-ignore"));
  }

  /**
   * Returns a click handler that scrolls a page to the specified element.
   */
  private scrollToHeader = (target: HTMLElement) => {
    return (event: React.SyntheticEvent) => {
      //decrement the history count to allow the return to previous page to work correctly
      const temp = this.state.historyCount - 1;
      this.setState({ historyCount: temp });
      event.preventDefault();
      document.location.hash = target.id;
      target.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
    };
  }

  /**
   * Creates a list of components to display from a list of links.
   * @param links
   */
  private renderLinks(links: Link[], listStyle: string): JSX.Element {
    // for each link render a <li> element with a link. If the link has got childNodes, additionaly render <ul> with child links.
    const elements = links.map((link, index) => {
      return (
        <li key={index}>
          <a onClick={this.scrollToHeader(link.element)} href={'#' + link.element.id}>{link.element.innerText}</a>
          {link.childNodes.length > 0 ? (<ul style={{ listStyleType: listStyle }}>{this.renderLinks(link.childNodes, listStyle)}</ul>) : ''}
        </li>
      );
    });

    return elements[0]; //Not sure what is going on here, but if the array is returned it renders twice while in 'Edit' mode.
  }

  /**
   * Force the component to re-render with a specified interval.
   * This is needed to get valid id values for headers to use in links. Right after the rendering headers won't have valid ids, they are assigned later once the whole page got rendered.
   * The component will display the correct list of headers on the first render and will be able to process clicks (as a link to an HTMLElement is stored by the component).
   * Once valid ids got assigned to headers by SharePoint code, the component will get valid ids for headers. This way a link from ToC can be copied by a user and it will be a valid link to a header.
   */
  public componentDidMount() {
    setInterval(() => {
      this.setState({});
    }, TableOfContents.timeout);
  }

  /**
   * Event for the back to previous page link. 
   * It uses the history count to work out how many pages to go back, as each click to a header results in history
   */
  public backToPreviousPage() {
    window.history.go(this.state.historyCount);
  }

  /**
   * Render the back to previous link
   */
  private renderBackToPreviousLink = (listStyle: string): JSX.Element => {
    if (this.props.showPreviousPageLink) {
      return (
        <div className={styles.backItem} ><ul style={{ listStyleType: listStyle }}><li><a href="#" onClick={() => this.backToPreviousPage()}>{this.props.previousPageText ? this.props.previousPageText : strings.previousPageDefaultValue}</a></li></ul></div>
      );
    }
    else {
      return null;
    }
  }

  /**
   * Modify the CSS of the appropriate HTML elements based on the wepart ID to enable sticky mode.
   * This does involve modifying HTML elements outside of the webpart, so may well break in the furture if Microsoft change their HTML\CSS etc.
   */
  private configureSticky() {

    const HTMLElementSticky: HTMLElement = document.querySelector("[id='" + this.props.webpartId + "']");
    if (HTMLElementSticky != null) {
      if (this.props.enableStickyMode) {

        HTMLElementSticky.classList.add(styles.sticky);
        HTMLElementSticky.parentElement.parentElement.classList.add(styles.height100pc);

      }
      else {

        HTMLElementSticky.classList.remove(styles.sticky);
        HTMLElementSticky.parentElement.parentElement.classList.remove(styles.height100pc);
      }
    }
  }

  public render(): JSX.Element {
    // get headers, then filter out empty and headers from <aside> tags
    const listStyle = escape(this.props.listStyle) === "default" ? "" : this.props.listStyle;
    const querySelector = this.getQuerySelector(this.props);
    const headers = this.getHtmlElements(querySelector).filter(this.filterEmpty).filter(this.filterAside).filter(this.filterTocIgnore);
    // create a list of links from headers
    const links = this.getLinks(headers);
    // create components from a list of links
    const toc = (<ul style={{ listStyleType: listStyle }}>{this.renderLinks(links, listStyle)}</ul>);
    // create previous page link
    const previousPage = (this.renderBackToPreviousLink(listStyle));
    // add CSS class to hide in mobile view if needed
    const hideInMobileViewClass = this.props.hideInMobileView ? (styles.hideInMobileView) : '';
    // add CSS class to hide title if requested
    const titleClass = this.props.hideTitle ? (styles.hideTitle) : "cke_editable h2 " + styles.title;
    // set title text
    const titleText = this.props.titleText ? this.props.titleText : strings.titleDefaultValue;
    // set Sticky parameters
    this.configureSticky();

    return (
      <section className={styles.tableOfContents}>
        <div className={hideInMobileViewClass}>
          <nav>
            <div className={titleClass}>
              <h2 data-toc-ignore="true">{escape(titleText)}</h2>
            </div>
            {toc}
            {previousPage}
          </nav>
        </div>
      </section>
    );
  }
}
