/**
 * @module botbuilder-dialogs-adaptive
 */
/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import { Expression, StringExpression, StringExpressionConverter } from 'adaptive-expressions';
import { Converter, ConverterFactory, DialogContext } from 'botbuilder-dialogs';
import { InputDialog, InputDialogConfiguration, InputState } from './inputDialog';

export interface TextInputConfiguration extends InputDialogConfiguration {
    outputFormat?: string | Expression | StringExpression;
}

export class TextInput extends InputDialog implements TextInputConfiguration {
    public static $kind = 'Microsoft.TextInput';

    public outputFormat: StringExpression;

    public getConverter(property: keyof TextInputConfiguration): Converter | ConverterFactory {
        switch (property) {
            case 'outputFormat':
                return new StringExpressionConverter();
            default:
                return super.getConverter(property);
        }
    }

    /**
     * @protected
     */
    protected onComputeId(): string {
        return `TextInput[${this.prompt && this.prompt.toString()}]`;
    }

    /**
     * @protected
     * Called when input has been received.
     * @param dc The [DialogContext](xref:botbuilder-dialogs.DialogContext) for the current turn of conversation.
     * @returns [InputState](xref:botbuilder-dialogs-adaptive.InputState) which reflects whether input was recognized as valid or not.
     */
    protected async onRecognizeInput(dc: DialogContext): Promise<InputState> {
        // Treat input as a string
        let input: string = dc.state.getValue(InputDialog.VALUE_PROPERTY).toString();

        if (this.outputFormat) {
            const value = this.outputFormat.getValue(dc.state);
            if (value && value.trim()) {
                input = value;
            }
        }

        // Save formated value and ensure length > 0
        dc.state.setValue(InputDialog.VALUE_PROPERTY, input);
        return input.length > 0 ? InputState.valid : InputState.unrecognized;
    }
}
