import React, {FunctionComponent, useState, useCallback} from "react";
import {render} from "react-dom";
//import { Stack, MessageBar, MessageBarType, IStackProps } from "@fluentui/react";
import {Stack, IStackProps} from "@fluentui/react/lib/Stack"
import {MessageBar, MessageBarType} from "@fluentui/react/lib/MessageBar"

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

  const MBError = (props: IMessageBarProps) => (
    <MessageBar
      messageBarType={MessageBarType.error}
      isMultiline={false}
      onDismiss={props.onDismiss}
      dismissButtonAriaLabel="Close"
    >
      {props.messageText}
    </MessageBar>
  );

  const BasicMessageBar: FunctionComponent<IMessageBarProps> = ({ messageText }) => {
    const [visible, setVisible] = useState<boolean | undefined>(true);

    const setVisibility = useCallback((visible: boolean) => setVisible(visible), []);

    const props: IMessageBarProps = { messageText, visible: true, onDismiss: () => setVisibility(false) };

    return visible !== false ? (
      <Stack {...horizontalStackProps}>
        <Stack {...verticalStackProps}>
          <MBError {...props} />
        </Stack>
      </Stack>
    ) : null;
  };

  render(<BasicMessageBar messageText={messageText} />, document.getElementById("messagebar"));
}
