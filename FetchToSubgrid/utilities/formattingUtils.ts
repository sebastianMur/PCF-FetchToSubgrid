import { IDataverseService } from '../services/dataverseService';

export const formatNumber = (_service: IDataverseService, value: string) =>
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore
  Number.parseLocale(value?.toString().split(' ')[0], _service.getContext().client.locale);

export const formatDateShort =
(_service: IDataverseService, value: Date, includeTime?: boolean): string =>
  _service.getContext().formatting.formatDateShort(value, includeTime);
