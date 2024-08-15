import * as React from 'react'; // Import React library for building components
import styles from './Marquee.module.scss'; // Import SCSS module for styling
import { IMarqueeProps } from './IMarqueeProps'; // Import the interface for the component props
import Marquee from 'react-fast-marquee'; // Import the Marquee component from the 'react-fast-marquee' library
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; // Import SPHttpClient to interact with SharePoint REST API

// Define the interface for the component's state
export interface IMarqueeState {
  items: any[]; // Array to hold list items retrieved from SharePoint
  fields: any[]; // Array to hold fields of the list items
  message: string; // Custom message to display in the marquee
}

// Define the main MarqueeComponent class which extends React.Component
export default class MarqueeComponent extends React.Component<IMarqueeProps, IMarqueeState> {
  constructor(props: IMarqueeProps) {
    super(props);
    // Initialize the state with empty arrays and a default message
    this.state = {
      items: [],
      fields: [],
      message: this.props.customMessage || 'Edit webpart to show custom message',
    };
  }

  // Lifecycle method that gets called after the component is mounted
  public componentDidMount(): void {
    this._getListItems(); // Fetch the list items and fields from SharePoint
  }

  // Private method to fetch list items and fields from SharePoint
  private _getListItems(): void {
    if (!this.props.selectedList) { // If no list is selected, do nothing
      return;
    }
    
    // Construct the API endpoint URLs for fetching list items and fields
    const listUrl = `${this.props.siteUrl}/_api/web/lists(guid'${this.props.selectedList}')/items?$select=*`;
    const fieldsUrl = `${this.props.siteUrl}/_api/web/lists(guid'${this.props.selectedList}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;

    // Fetch the fields from the SharePoint list
    this.props.spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: any) => {
        // Filter out system fields and map the remaining fields
        const systemFields = ["ID", "ContentType", "Modified", "Created", "Author", "Editor", "Attachments", "ContentTypeId", "Title"];
        const fields = data.value
          .filter((field: any) => !systemFields.includes(field.InternalName))
          .map((field: any) => ({ title: field.Title, internalName: field.InternalName }));
        console.log('Fields:', fields); // Log the fields to the console
        this.setState({ fields }); // Update the state with the filtered fields

        // Fetch the list items using the previously constructed URL
        return this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => response.json())
          .then((data: any) => {
            let items = data.value; // Get the list items from the response
            console.log('Items:', items); // Log the items to the console
            if (this.props.randomize) {
              items = this._shuffleArray(items); // Randomize the items if specified
            }
            this.setState({ items }); // Update the state with the items
          });
      });
  }

  // Private method to shuffle an array (Fisher-Yates algorithm)
  private _shuffleArray(array: any[]): any[] {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]]; // Swap elements
    }
    return array; // Return the shuffled array
  }

  // Render method to output the JSX of the component
  public render(): React.ReactElement<IMarqueeProps> {
    const { fields, items, message } = this.state; // Destructure state properties
    const { showFieldLabels, showCustomMessage, headerColor, customMessageColor, customMessageBold, imageUrl } = this.props; // Destructure props

    // Define custom styles for the custom message
    const customMessageStyle = {
      color: customMessageColor,
      fontWeight: customMessageBold ? 'bold' : 'normal'
    };

    // Create an array of content elements to display in the marquee
    const content = [
      showCustomMessage && (
        <div key="message" className={styles.marqueeItem}>
          <p style={customMessageStyle}>{message}</p> {/* Display the custom message */}
        </div>
      ),
      ...items.map((item, index) => (
        <div key={index} className={styles.marqueeItem}>
          {fields.map((field, idx) => (
            item[field.internalName] && (
              <p key={idx}>
                {showFieldLabels && <strong style={{ color: headerColor }}>{field.title}:</strong>} {item[field.internalName]}
              </p> // Display each field's label and value
            )
          ))}
        </div>
      ))
    ];

    // Return the final JSX structure including the marquee and optional image
    return (
      <section className={styles.marquee}>
        {imageUrl && (
          <div className={styles.welcomeImageContainer}>
            <img src={imageUrl} alt="Uploaded image" className={styles.welcomeImage} /> {/* Display the uploaded image if provided */}
          </div>
        )}
        <Marquee pauseOnHover={true} delay={2} gradient={true} gradientColor="white" gradientWidth={10}>
          {content} {/* Display the content inside the marquee */}
        </Marquee>
      </section>
    );
  }
}
