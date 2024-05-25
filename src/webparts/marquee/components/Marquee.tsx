import * as React from 'react';
import styles from './Marquee.module.scss';
import { IMarqueeProps } from './IMarqueeProps';
import Marquee from 'react-fast-marquee';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IMarqueeState {
  items: any[];
  fields: any[];
  message: string;
}

export default class MarqueeComponent extends React.Component<IMarqueeProps, IMarqueeState> {
  constructor(props: IMarqueeProps) {
    super(props);
    this.state = {
      items: [],
      fields: [],
      message: this.props.customMessage || 'Edit webpart to show custom message',
    };
  }

  public componentDidMount(): void {
    this._getListItems();
  }

  private _getListItems(): void {
    if (!this.props.selectedList) {
      return;
    }
    const listUrl = `${this.props.siteUrl}/_api/web/lists(guid'${this.props.selectedList}')/items?$select=*`;
    const fieldsUrl = `${this.props.siteUrl}/_api/web/lists(guid'${this.props.selectedList}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`;

    this.props.spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: any) => {
        const systemFields = ["ID", "ContentType", "Modified", "Created", "Author", "Editor", "Attachments", "ContentTypeId", "Title"];
        const fields = data.value
          .filter((field: any) => !systemFields.includes(field.InternalName))
          .map((field: any) => ({ title: field.Title, internalName: field.InternalName }));
        console.log('Fields:', fields);
        this.setState({ fields });

        return this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => response.json())
          .then((data: any) => {
            let items = data.value;
            console.log('Items:', items);
            if (this.props.randomize) {
              items = this._shuffleArray(items);
            }
            this.setState({ items });
          });
      });
  }

  private _shuffleArray(array: any[]): any[] {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }

  public render(): React.ReactElement<IMarqueeProps> {
    const { fields, items, message } = this.state;
    const { showFieldLabels, showCustomMessage, headerColor, customMessageColor, customMessageBold } = this.props;

    const customMessageStyle = {
      color: customMessageColor,
      fontWeight: customMessageBold ? 'bold' : 'normal'
    };

    const content = [
      showCustomMessage && (
        <div key="message" className={styles.marqueeItem}>
          <p style={customMessageStyle}>{message}</p>
        </div>
      ),
      ...items.map((item, index) => (
        <div key={index} className={styles.marqueeItem}>
          {fields.map((field, idx) => (
            item[field.internalName] && (
              <p key={idx}>
                {showFieldLabels && <strong style={{ color: headerColor }}>{field.title}:</strong>} {item[field.internalName]}
              </p>
            )
          ))}
        </div>
      ))
    ];

    return (
      <section className={styles.marquee}>
        <Marquee pauseOnHover={true} delay={2} gradient={true} gradientColor="white" gradientWidth={10}>
          {content}
        </Marquee>
      </section>
    );
  }
}
