import * as React from 'react';
import { useState } from 'react';
import {
  Callout,
  ComboBox,
  DefaultButton,
  DirectionalHint,
  IColumn,
  IComboBox,
  IComboBoxOption,
  IconButton,
  PrimaryButton,
} from '@fluentui/react';
import { filterMenuStyle } from '../styles/componentsStyles';
import { FilterInputComponent } from './FilterInputComponent';
import { getOptionsByAttributeType, isInputHidden } from '../utilities/filterUtils';
import { IDataverseService } from '../services/dataverseService';
import { IInitialInputValue } from './List';

interface FilterPopupProps {
  dataverseService: IDataverseService;
  column: IColumn;
  target: HTMLElement;
  initialInputValues: IInitialInputValue[];
  hideFilterMenu: () => void;
  handleConfirm: (selectedOption: IComboBoxOption, inputValue: string) => void;
}

export const FilterPopup: React.FC<FilterPopupProps> = ({
  dataverseService,
  column,
  target,
  initialInputValues,
  hideFilterMenu,
  handleConfirm,
}) => {
  const checkBoxOptions = getOptionsByAttributeType(column.data.attributeType);

  const matchingValue = initialInputValues.find(val =>
    val.fieldName === column.fieldName,
  );
  const isHidden = !isInputHidden(matchingValue?.selectedOption?.key as number || 0);

  const [selectedOption, setSelectedOption] = useState<IComboBoxOption>(
    matchingValue?.selectedOption || checkBoxOptions[0]);
  const [inputValue, setInputValue] = useState('');
  const [showInput, setShowInput] = useState(isHidden);
  const [isButtonDisable, setDisableButton] = useState(false);
  const [inputErrorMessage, setInputErrorMessage] = useState('');

  const onOptionChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption | undefined,
  ) => {
    setInputValue('');
    isInputHidden(option?.key as number || 0)
      ? setShowInput(false) : setShowInput(true);

    if (option) {
      setSelectedOption(option);
    }

    setDisableButton(false);
  };

  const handleApplyFilter = () => {
    handleConfirm(selectedOption!, inputValue);
    hideFilterMenu();
  };

  const onTextChange = (value: string) => {
    setInputValue(value);
  };

  return (
    <>
      <Callout
        gapSpace={2}
        isBeakVisible={false}
        target={target}
        directionalHint={DirectionalHint.bottomLeftEdge}
        onDismiss={hideFilterMenu}
        style={filterMenuStyle}
      >
        <div className='filterTitleContainer'>
          <h3 className='filterTitle'>Filter By</h3>
          <IconButton
            className='filterCloseButton'
            iconProps={{ iconName: 'Cancel' }}
            onClick={hideFilterMenu}
          />
        </div>
        <div className='filterSection'>
          <ComboBox
            className='filterOptions'
            defaultSelectedKey={selectedOption.key}
            options={checkBoxOptions}
            onChange={onOptionChange}
          />
          {showInput &&
          <FilterInputComponent
            dataverseService={dataverseService}
            column={column}
            inputValue={inputValue}
            selectedOption={selectedOption}
            setInputValue={setInputValue}
            onTextChange={e => onTextChange(e.currentTarget.value)}
            initialInputValues={initialInputValues}
            inputErrorMessage={inputErrorMessage}
            setInputErrorMessage={setInputErrorMessage}
          />
          }

          <div className='filterButtonContainer'>
            <PrimaryButton
              text='Apply'
              className='filterApplyButton'
              onClick={handleApplyFilter}
              disabled={isButtonDisable}
            />
            <DefaultButton
              text='Clear'
              disabled={isButtonDisable}
              onClick={() => { setInputValue(''); }}
            />
          </div>
        </div>
      </Callout>
    </>
  );
};
