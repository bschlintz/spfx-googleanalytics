import * as React from 'react';
import styles from './AnalyticsSettingsButton.module.scss';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Button, PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

export class IAnalyticsSettingsButtonProps {
  public googleTrackingId: string;
  public onSaveFunction: Function;
}

export class AnalyticsSettingsButton extends React.Component<IAnalyticsSettingsButtonProps, { showModal: boolean; }> {
  private _textGoogleAnalyticsTrackingId: any;

  constructor(props) {
    super(props);
    this.state = {
      showModal: false
    };
  }
  public render(): React.ReactElement<any> {
    return (
      <div className={styles.container}>
        <IconButton
          iconProps={{ iconName: "BarChart4", className: styles.iconButton }}
          onClick={this._showModal}
          title="Google Analytics Settings"
          ariaLabel="Google Analytics Settings"
        />
        <Dialog
          isOpen={this.state.showModal}
          onDismiss={this._closeModal}
          isBlocking={false}
          title={"Google Analytics Settings"}
          subText={"Update Google Analytics Tracking ID used by this site."}
        >
          <TextField
            label="Tracking ID"
            defaultValue={this.props.googleTrackingId}
            componentRef={r => this._textGoogleAnalyticsTrackingId = r}
          ></TextField>
          <DialogFooter>
            <PrimaryButton onClick={this._onClickSave} text="Save" />
            <Button onClick={this._closeModal} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _onClickSave = async (): Promise<void> => {
    try {
      await this.props.onSaveFunction(this._textGoogleAnalyticsTrackingId.value);
      //refresh page on save
      location.reload();

      // this.setState({ showModal: false });
    }
    catch (error) {
      console.log("ERROR: Unable to save Google Analytics Tracking Id", error);
      this.setState({ showModal: false });
    }
  }

  private _showModal = (): void => {
    this.setState({ showModal: true });
  }

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  }
}
