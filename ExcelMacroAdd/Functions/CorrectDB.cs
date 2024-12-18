﻿using ExcelMacroAdd.BisinnesLayer.Interfaces;
using ExcelMacroAdd.DataLayer.Entity;
using ExcelMacroAdd.Serializable.Entity.Interfaces;
using ExcelMacroAdd.UserException;
using System;
using System.Data;
using System.Windows.Forms;

namespace ExcelMacroAdd.Functions
{
    internal sealed class CorrectDb : AbstractFunctions
    {
        private readonly IJournalData accessData;
        private readonly IFillingOutThePassportSettings resources;

        public CorrectDb(IJournalData accessData, IFillingOutThePassportSettings resources)
        {
            this.accessData = accessData;
            this.resources = resources;
        }

        public override async void Start()
        {

            if (Application.ActiveWorkbook.Name != resources.NameFileJournal) // Проверка по имени книги
            {
                MessageWarning("Функция работает только в \"Журнале учета НКУ\" текущего года. \n Пожайлуста откройте необходимую книгу Excel.",
                     "Имя книги не совпадает с целевой");
                return;
            }
            DialogResult dialogResult = MessageBox.Show(@"Вы уверены, что хотите изменить запись в БД? Пожалуйста будте очень внимательны, изменения коснуться всех пользователей.",
                                                        @"Контрольный вопрос", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                var firstRow = Cell.Row; // Вычисляем верхний элемент
                string sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);

                try
                {
                    if (!(await accessData.AccessJournalNku.GetEntityJournal(sArticle.ToLower()) is BoxBase journalNku))
                    {
                        MessageWarning($"В базе данных такого артикула нет.\n Необходимо сначала его занести. \nАртикул = {sArticle}",
                        "Ошибка записи!");
                        return;
                    }

                    int.TryParse(Convert.ToString(Worksheet.Cells[firstRow, 11].Value2), out int sIp);

                    string sClimate = Convert.ToString(Worksheet.Cells[firstRow, 12].Value2);
                    string sWeight = Convert.ToString(Worksheet.Cells[firstRow, 13].Value2);
                    string sHeight = Convert.ToString(Worksheet.Cells[firstRow, 14].Value2);
                    string sWidth = Convert.ToString(Worksheet.Cells[firstRow, 15].Value2);
                    string sDepth = Convert.ToString(Worksheet.Cells[firstRow, 16].Value2);
                    sArticle = Convert.ToString(Worksheet.Cells[firstRow, 26].Value2);
                    string sMaterial = Convert.ToString(Worksheet.Cells[firstRow, 29].Value2);
                    string sExecution = Convert.ToString(Worksheet.Cells[firstRow, 30].Value2);

                    if (sHeight == null || sWidth == null || sDepth == null || sArticle == null || sMaterial == null || sExecution == null)
                    {
                        MessageWarning($"Одно из обязательных полей не заполнено. Пожайлуста запоните все поля и еще раз повторрите запись. \n Артикул = {sArticle}",
                            "Ошибка записи");
                        return;
                    }

                    var materialEntity = await accessData.AccessJournalNku.GetMaterialEntityByName(sMaterial) ?? throw new DataBaseNotFoundValueException($"Введенный материал шкафа \"{sMaterial}\" недопустим, пожайлуста используйте значение \"Пластик\" или \"Металл\"");
                    var executionEntity = await accessData.AccessJournalNku.GetExecutionEntityByName(sExecution) ?? throw new DataBaseNotFoundValueException($"Введенное исполнение шкафа \"{sExecution}\" недопустимо, пожайлуста используйте значение \"напольное\", или \"навесное\", или \"встраиваемое\", или \"навесное для IT оборудования\", или \"напольное для IT оборудования\".");

                    journalNku.Ip = sIp;
                    journalNku.Climate = sClimate == "-" ? null : sClimate;
                    journalNku.Weight = sWeight == "-" ? null : sWeight;
                    journalNku.Height = sHeight;
                    journalNku.Width = sWidth;
                    journalNku.Depth = sDepth;
                    journalNku.Article = sArticle.ToLower();
                    journalNku.MaterialBoxId = materialEntity.Id;
                    journalNku.ExecutionBoxId = executionEntity.Id;

                    accessData.AccessJournalNku.WriteUpdateDb(journalNku);

                    MessageInformation($"Запись успешно изменена! \nПоздравляем! \nАртикул = {sArticle}",
                                "Запись успешна!");
                }

                catch (DataBaseNotFoundValueException e)
                {
                    MessageError($"Произошла ошибка, скорее всего непавильно было указано одно из значений. {e.Message}",
                        "Ошибка базы данных");
                }

                catch (DataException)
                {
                    MessageError("Не удалось подключиться к базе данных, просьба проверить наличие или доступность файла базы данных",
                        "Ошибка базы данных");
                }
                catch (Exception e)
                {
                    MessageError($"Произошла непредвиденная ошибка, пожайлуста сделайте скриншот ошибки, и передайте его разработчику.\n {e.Message}",
                        "Ошибка базы данных");
                }
            }
        }
    }
}
