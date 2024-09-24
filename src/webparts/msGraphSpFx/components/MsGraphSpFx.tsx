import * as React from 'react';
import type { IMsGraphSpFxProps } from './IMsGraphSpFxProps';

import { IMsGraphSpFxState } from './IMsGraphSpFxState';

import { GraphError, ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  Persona,
  PersonaSize
  // } from 'office-ui-fabric-react/lib/components/Persona';
} from '@fluentui/react'

// import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { Link } from '@fluentui/react'

export default class MsGraphSpFx extends React.Component<IMsGraphSpFxProps, IMsGraphSpFxState> {

  constructor(props: IMsGraphSpFxProps) {
    super(props);
    this.state = {
      name: '',
      email: '',
      phone: '',
      image: undefined
    };
  }

  public componentDidMount(): void {
    /* eslint-disable @typescript-eslint/no-floating-promises */
    this.props.graphClient
      .api('me')
      .get((error: GraphError, user: MicrosoftGraph.User) => {
        this.setState({
          name: user.displayName as string,
          email: user.mail as string,
          phone: user.businessPhones?.[0]
        });
      });

    this.props.graphClient
      .api('/me/photo/$value')
      .responseType(ResponseType.BLOB)
      .get((error: GraphError, photoResponse: Blob) => {
        const blobUrl = window.URL.createObjectURL(photoResponse);
        this.setState({ image: blobUrl });
      });
    /* eslint-enable @typescript-eslint/no-floating-promises */
  }

  public render(): React.ReactElement<IMsGraphSpFxProps> {
    return (
      <Persona primaryText={this.state.name}
        secondaryText={this.state.email}
        onRenderSecondaryText={this._renderMail}
        tertiaryText={this.state.phone}
        onRenderTertiaryText={this._renderPhone}
        imageUrl={this.state.image}
        size={PersonaSize.size100} />
    );
  }

  private _renderMail = (): JSX.Element => {
    if (this.state.email) {
      return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
    } else {
      return <div />;
    }
  }

  private _renderPhone = (): JSX.Element => {
    if (this.state.phone) {
      return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>;
    } else {
      return <div />;
    }
  }
}
