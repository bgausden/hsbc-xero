import React from "react";
import {
  Stack,
  MessageBar,
  MessageBarType,
  IStackProps
} from "office-ui-fabric-react";
import ReactDOM from "react-dom";

/* global document */

export function displayMessageBar(messageText: string) {
  interface IMessageBarProps {
    messageText?: string;
    visible?: boolean;
    onDismiss?: () => void;
  }

  const horizontalStackProps: IStackProps = {
    horizontal: true,
    tokens: { childrenGap: 16 }
  };
  const verticalStackProps: IStackProps = {
    styles: { root: { overflow: "hidden", width: "100%" } },
    tokens: { childrenGap: 20 }
  };

  const ErrorExample = (props: IMessageBarProps) => (
    <MessageBar
      messageBarType={MessageBarType.error}
      isMultiline={false}
      onDismiss={props.onDismiss}
      dismissButtonAriaLabel="Close"
    >
      {props.messageText}
    </MessageBar>
  );

  const BasicMessageBar: React.FunctionComponent<IMessageBarProps> = ({ messageText }) => {
    const [visible, setVisible] = React.useState<boolean | undefined>(true);

    const setVisibility = React.useCallback((visible: boolean) => setVisible(visible), []);

    const props: IMessageBarProps = { messageText, visible: true, onDismiss: () => setVisibility(false) };

    return visible !== false ? (
      <Stack {...horizontalStackProps}>
        <Stack {...verticalStackProps}>
          <ErrorExample {...props} />
        </Stack>
      </Stack>
    ) : null;
  };

  ReactDOM.render(<BasicMessageBar messageText={messageText} />, document.getElementById("messagebar"));
}
