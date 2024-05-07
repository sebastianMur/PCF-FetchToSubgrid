export const getDateFormatWithHyphen = (date: Date | undefined) => {
  if (date === undefined) return '';

  const day = date.getDate() > 9 ? date.getDate() : `0${date.getDate()}`;
  const month = date.getMonth() + 1 > 9 ? `${date.getMonth() + 1}` : `0${date.getMonth() + 1}`;

  return `${date.getFullYear()}-${month}-${day}`;
};
