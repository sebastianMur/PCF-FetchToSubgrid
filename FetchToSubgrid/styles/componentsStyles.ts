import { IComboBoxStyles } from '@fluentui/react';

export const filterMenuStyle = {
  maxWidth: '280px',
};

export const wholeFormatStyles = (required: boolean): Partial<IComboBoxStyles> => ({
  optionsContainer: {
    maxHeight: 260,
  },
  container: {
    marginRight: required ? '10px' : '0px',
  },
});
