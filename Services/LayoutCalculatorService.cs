using System;
using System.Collections.Generic;
using System.Linq;
using TocBuilder_dotnet_framework.Models;

namespace TocBuilder_dotnet_framework.Services
{
    public static class LayoutConstants
    {
        public const float TitleHeight = 120f;
        public const float CaptionHeight = 20f;
        public const float BottomMargin = 50f;
        public const float MinThumbWidth = 80f;
        public const float MinThumbHeight = 60f;

        public const float DefaultSlideWidth = 960f;   // 13.333 in × 72 = 960 (16:9)
        public const float DefaultSlideHeight = 540f;  // 7.5 in × 72 = 540 (16:9)
    }

    public static class LayoutCalculatorService
    {
        //public const float SLIDE_WIDTH_POINTS = 914.4f; // Пример для 4:3
        //public const float SLIDE_HEIGHT_POINTS = 685.8f; // Пример для 4:3

        public static (int Columns, float ThumbWidth, float ThumbHeight, float RowHeight) CalculateOptimalLayout(
            int slideCount,
            float margin,
            int desiredColumns = -1,
            float slideWidth = LayoutConstants.DefaultSlideWidth,
            float slideHeight = LayoutConstants.DefaultSlideHeight)
        {
            float slideAspectRatio = slideWidth / slideHeight;
            const float titleHeight = LayoutConstants.TitleHeight;
            const float bottomMargin = LayoutConstants.BottomMargin;
            const float captionHeight = LayoutConstants.CaptionHeight;
            const float minThumbWidth = LayoutConstants.MinThumbWidth;
            const float minThumbHeight = LayoutConstants.MinThumbHeight;


            // Доступная высота на мирниатюры и подписи
            float availableHeight = slideHeight - titleHeight - bottomMargin;
            float availableWidth = slideWidth - (margin * 2);

            // Вспомогательная функция для проверки конфигурации
            (bool Fits, float ThumbW, float ThumbH, float RowH) TryColumns(int cols)
            {
                // Ширина каждой миниатюры при `cols` колонках
                float thumbW = (availableWidth - margin * (cols - 1)) / cols;
                float thumbH = thumbW / slideAspectRatio;

                int rows = (int)Math.Ceiling((double)slideCount / cols);
                float rowH = thumbH + margin + captionHeight;

                float totalH = rows * rowH; // Общая высота сетки
                bool fits = totalH <= availableHeight && thumbW >= minThumbWidth && thumbH >= minThumbHeight;

                // Если не помещается по высоте — попробуем сжать
                if (!fits && totalH > availableHeight)
                {
                    // Расчёт максимальной допустимой высоты под миниатюры+подписи
                    // Доступная высота на миниатюры и подписи (без учёта межстрочных отступов)
                    float maxThumbAndCaptionHeight = availableHeight / rows;
                    // Вычитаем captionHeight и margin между строками (margin на каждую строку, кроме последней)
                    // Для упрощения — на строку: thumbH + captionHeight + margin (кроме последней строки)
                    // Но чтобы не усложнять — используем приближение: усреднённая высота на строку = totalH/rows
                    // Альтернатива: решить уравнение: rows * (thumbH + captionHeight) + (rows - 1) * margin ≤ availableHeight
                    // => rows * thumbH ≤ availableHeight - rows * captionHeight - (rows - 1) * margin
                    float maxThumbH = (availableHeight - rows * captionHeight - (rows - 1) * margin) / rows;

                    if (maxThumbH > 0)
                    {
                        thumbH = Math.Max(minThumbHeight, maxThumbH);
                        thumbW = thumbH * slideAspectRatio; // сохраняем соотношение

                        // Проверим, умещается ли ширина
                        if (thumbW > (availableWidth - margin * (cols - 1)) / cols)
                        {
                            // Ширина превышает — сжимаем по ширине
                            thumbW = (availableWidth - margin * (cols - 1)) / cols;
                            thumbH = thumbW / slideAspectRatio;
                        }

                        rowH = thumbH + margin + captionHeight;
                        totalH = rows * rowH;
                        fits = totalH <= availableHeight && thumbW >= minThumbWidth && thumbH >= minThumbHeight;
                    }
                }

                return (fits, thumbW, thumbH, rowH);
            }

            // 1. Если задано желаемое количество колонок — попробуем его сначала
            if (desiredColumns > 0)
            {
                var result = TryColumns(desiredColumns);
                if (result.Fits)
                    return (desiredColumns, result.ThumbW, result.ThumbH, result.RowH);
            }

            // 2. Иначе — перебор от 1 до 8 колонок, ищем **наибольшую площадь миниатюры**
            (int bestCols, float bestW, float bestH, float bestRH) = (1, 0, 0, 0);
            float bestArea = 0;

            for (int cols = 1; cols <= 8; cols++)
            {
                var (fits, w, h, rh) = TryColumns(cols);
                if (fits)
                {
                    float area = w * h;
                    if (area > bestArea)
                    {
                        bestArea = area;
                        (bestCols, bestW, bestH, bestRH) = (cols, w, h, rh);
                    }
                }
            }

            // 3. Если ничего не подошло — используем fallback (1 колонка, принудительное сжатие)
            if (bestArea == 0)
            {
                var fallback = TryColumns(1);
                if (fallback.Fits)
                    return (1, fallback.ThumbW, fallback.ThumbH, fallback.RowH);

                // Крайний fallback: жёстко используем минимальные размеры
                float w = minThumbWidth;
                float h = Math.Max(minThumbHeight, minThumbWidth / slideAspectRatio); // сохраняем пропорции
                float rh = h + margin + captionHeight;
                return (Math.Min(4, slideCount), w, h, rh);
            }

            return (bestCols, bestW, bestH, bestRH);
        }

        public static List<PreviewItem> GeneratePreviewItems(
            List<SlideItem> selectedSlides,
            int columns,
            int margin,
            float slideWidth = LayoutConstants.DefaultSlideWidth,
            float slideHeight = LayoutConstants.DefaultSlideHeight)
        {
            var previewItems = new List<PreviewItem>();

            if (selectedSlides == null || !selectedSlides.Any()) return previewItems;

            var layoutInfo = CalculateOptimalLayout(
                selectedSlides.Count,
                margin,
                columns,
                slideWidth,
                slideHeight);

            float thumbWidth = layoutInfo.ThumbWidth;
            float thumbHeight = layoutInfo.ThumbHeight;
            float rowHeight = layoutInfo.RowHeight;

            const float titleHeight = LayoutConstants.TitleHeight;
            const float yStart = titleHeight;

            for (int i = 0; i < selectedSlides.Count; i++)
            {
                int row = i / layoutInfo.Columns;
                int col = i % layoutInfo.Columns;

                float x = margin + col * (thumbWidth + margin);
                float y = yStart + row * rowHeight;

                previewItems.Add(new PreviewItem
                {
                    X = x,
                    Y = y,
                    Width = thumbWidth,
                    Height = thumbHeight,
                    Caption = $"Слайд {selectedSlides[i].Number}",
                    Thumbnail = selectedSlides[i].Thumbnail
                });
            }

            return previewItems;
        }

        public static (double Width, double Height) CalculateCanvasSize(List<PreviewItem> previewItems)
        {
            if (previewItems == null || !previewItems.Any())
            {
                return (LayoutConstants.DefaultSlideWidth, LayoutConstants.DefaultSlideHeight); // Возвращаем размер слайда по умолчанию, если нет элементов
            }

            // Находим максимальные координаты
            double maxX = previewItems.Max(item => item.Right);
            double maxY = previewItems.Max(item => item.Bottom);

            // Добавляем небольшой запас (50 точек) снизу и справа
            double padding = 50;
            double canvasWidth = maxX + padding;
            double canvasHeight = maxY + padding;

            // Убедимся, что размеры не меньше стандартного размера слайда
            canvasWidth = Math.Max(canvasWidth, LayoutConstants.DefaultSlideWidth);
            canvasHeight = Math.Max(canvasHeight, LayoutConstants.DefaultSlideHeight);

            return (canvasWidth, canvasHeight);
        }
    }
}