using DaEtoZhe.Desktop.Controls;
using DaEtoZhe.Desktop.Models;   // <-- ВАЖНО: для AppSettings
// Остальные using
using DaEtoZhe.Desktop.Services; // <-- ВАЖНО: для методов расширения (GetWasherId/SetWasherId)
using DaEtoZhe.Desktop.ViewModels;
using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ContractsBranch = DaEtoZhe.Contracts.Models.Branch;
using ContractsClient = DaEtoZhe.Contracts.Models.Client;
// === ЯВНЫЕ АЛИАСЫ ДЛЯ КОНТРАКТНЫХ МОДЕЛЕЙ ===
using ContractsOrder = DaEtoZhe.Contracts.Models.Order;
using ContractsService = DaEtoZhe.Contracts.Models.Service;
using ContractsShift = DaEtoZhe.Contracts.Models.Shift;
using ContractsUser = DaEtoZhe.Contracts.Models.User;

namespace DaEtoZhe.Desktop
{
    public partial class AddEditOrderWindow : Window
    {
        private readonly ApiService _apiService;              // Для работы с сервером
        private readonly ContractsShift _currentShift;
        private readonly AddEditOrderViewModel _viewModel;

        private List<ContractsBranch> _allBranches; // Кэш филиалов на уровне класса

        public AddEditOrderWindow(
            AddEditOrderViewModel viewModel,
            ContractsShift currentShift = null,
            ContractsOrder order = null)
        {
            InitializeComponent();

            _apiService = new ApiService();
            _currentShift = currentShift;
            _viewModel = viewModel;

            _viewModel.Initialize(currentShift, order);
            DataContext = _viewModel;

            if (_viewModel.CurrentOrder != null &&
                !string.IsNullOrWhiteSpace(_viewModel.CurrentOrder.Department))
            {
                _viewModel.CurrentDepartment = _viewModel.CurrentOrder.Department;
            }
            else
            {
                _viewModel.CurrentDepartment = "Wash";
            }

            _ = LoadDictionariesAsync(order);
        }

        private static void SetComboBoxSource(
    CustomComboBox comboBox,
    IEnumerable<object> source,
    string displayMemberPath,
    string selectedValuePath)
        {
            if (comboBox == null)
                return;

            comboBox.DisplayMemberPath = displayMemberPath;
            comboBox.SelectedValuePath = selectedValuePath;
            comboBox.ItemsSource = source;
        }

        // Обработчик: при смене департамента обновляем список зон
        private void DepartmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
                return;

            if (_viewModel?.CurrentOrder == null)
                return;

            var department = DepartmentComboBox.SelectedValue as string;

            if (string.IsNullOrWhiteSpace(department))
                department = "Wash";

            _viewModel.CurrentDepartment = department;
            _viewModel.CurrentOrder.Department = department;

            RefreshZonesForSelectedBranch(_viewModel.CurrentOrder);
        }

        

        // Событие: если администратор переключил филиал ручками
        private void BranchComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
                return;

            if (_viewModel?.CurrentOrder == null)
                return;

            if (BranchComboBox.SelectedValue is int branchId && branchId > 0)
            {
                _viewModel.CurrentOrder.BranchId = branchId;
            }

            RefreshZonesForSelectedBranch(null);
        }

        private static List<ZoneItem> BuildZones(ContractsBranch branch)
        {
            var zones = new List<ZoneItem>();

            if (branch == null)
                return zones;

            for (int i = 1; i <= branch.WashBaysCount; i++)
            {
                zones.Add(new ZoneItem
                {
                    Name = $"🧽 Бокс {i} (Мойка)",
                    BoxNumber = i,
                    Department = "Wash"
                });
            }

            for (int i = 1; i <= branch.ServiceLiftsCount; i++)
            {
                zones.Add(new ZoneItem
                {
                    Name = $"🔧 Подъемник {i} (Сервис)",
                    BoxNumber = i,
                    Department = "Service"
                });
            }

            return zones;
        }

        private bool TryGetSelectedBranchId(out int branchId)
        {
            if (BranchComboBox?.SelectedValue is int id && id > 0)
            {
                branchId = id;
                return true;
            }

            branchId = _currentShift?.BranchId ?? AppSettings.CurrentBranchId;
            return branchId > 0;
        }

        private void RefreshZonesForSelectedBranch(ContractsOrder orderToRestore)
        {
            if (_viewModel?.CurrentOrder == null)
                return;

            if (!TryGetSelectedBranchId(out int branchId))
                return;

            var currentBranch = _allBranches?.FirstOrDefault(b => b.Id == branchId);
            var allZones = BuildZones(currentBranch);

            _viewModel.UpdateAvailableZones(allZones);
            SelectBestZone(orderToRestore);
        }

        private void SelectBestZone(ContractsOrder order)
        {
            if (_viewModel?.CurrentOrder == null)
                return;

            if (_viewModel.AvailableZones == null || !_viewModel.AvailableZones.Any())
            {
                if (ZoneComboBox != null)
                {
                    ZoneComboBox.SelectedIndex = -1;
                }

                _viewModel.CurrentOrder.BoxNumber = 0;
                return;
            }

            var currentDepartment = _viewModel.CurrentOrder.Department
                                    ?? _viewModel.CurrentDepartment
                                    ?? "Wash";

            if (order != null && order.BoxNumber > 0)
            {
                var existingZone = _viewModel.AvailableZones.FirstOrDefault(z =>
                    z.BoxNumber == order.BoxNumber &&
                    string.Equals(z.Department, order.Department, StringComparison.OrdinalIgnoreCase));

                if (existingZone != null)
                {
                    ZoneComboBox.SelectedValue = existingZone.BoxNumber;
                    _viewModel.CurrentOrder.BoxNumber = existingZone.BoxNumber;
                    _viewModel.CurrentOrder.Department = existingZone.Department;
                    return;
                }
            }

            var preferredZone = _viewModel.AvailableZones.FirstOrDefault(z =>
                string.Equals(z.Department, currentDepartment, StringComparison.OrdinalIgnoreCase));

            var zone = preferredZone ?? _viewModel.AvailableZones.First();

            ZoneComboBox.SelectedValue = zone.BoxNumber;
            _viewModel.CurrentOrder.BoxNumber = zone.BoxNumber;
            _viewModel.CurrentOrder.Department = zone.Department;
        }

        private async Task LoadDictionariesAsync(ContractsOrder order)
        {
            try
            {
                if (_viewModel?.CurrentOrder == null)
                    return;

                var allUsers = await _apiService.GetUsersAsync();
                var allClients = await _apiService.GetClientsAsync();
                _allBranches = await _apiService.GetBranchesAsync();

                int currentBranchId;

                if (order != null && order.BranchId > 0)
                {
                    currentBranchId = order.BranchId;
                }
                else if (_currentShift != null)
                {
                    currentBranchId = _currentShift.BranchId;
                }
                else
                {
                    currentBranchId = AppSettings.CurrentBranchId;
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] Loading dictionaries for BranchId: {currentBranchId}");

                var carCategories = await _apiService.GetCarCategoriesAsync(currentBranchId);
                SetComboBoxSource(BodyTypeComboBox, carCategories, "Name", "Id");

                var paymentMethods = await _apiService.GetPaymentMethodsAsync(currentBranchId);
                SetComboBoxSource(PaymentMethodComboBox, paymentMethods, "Name", "Name");

                var orderStatuses = await _apiService.GetOrderStatusesAsync(currentBranchId);
                SetComboBoxSource(StatusComboBox, orderStatuses, "DisplayName", "Name");

                // Мойщики теперь приходят через ViewModel и биндинг ItemsSource="{Binding Washers}"
                _viewModel.Washers = allUsers;

                // Клиенты пока грузим прямо в контрол, потому что во ViewModel нет свойства Clients
                ClientComboBox.ItemsSource = allClients;

                SetComboBoxSource(BranchComboBox, _allBranches, "Name", "Id");

                if (order != null && order.Id > 0)
                {
                    _ = _viewModel.LoadOrderExpensesAsync(order.Id);
                    _ = _viewModel.LoadOrderTimelineAsync(order.Id);
                }

                if (order != null)
                {
                    if (order.IsAppointment)
                    {
                        DurationTextBox.Text = order.DurationMinutes > 0
                            ? order.DurationMinutes.ToString()
                            : "60";
                    }

                    // Клиент: ставим выбранного клиента во ViewModel.
                    // SelectedItem="{Binding SelectedClient, Mode=TwoWay}" сам选ернёт комбобокс.
                    if (order.ClientId.HasValue)
                    {
                        var client = allClients?.FirstOrDefault(c => c.Id == order.ClientId.Value);

                        if (client != null)
                        {
                            _viewModel.SelectedClient = client;
                        }
                    }

                    // Сотрудник: SelectedValue биндится на SelectedWasherId.
                    if (order.GetWasherId() > 0)
                    {
                        _viewModel.SelectedWasherId = order.GetWasherId();
                    }

                    // Категория кузова: биндится на SelectedBodyTypeCategory.
                    if (order.BodyTypeCategory > 0)
                    {
                        _viewModel.SelectedBodyTypeCategory = order.BodyTypeCategory;
                    }

                    // Статус и оплата: можно оставить прямую установку SelectedValue,
                    // потому что CurrentOrder — обычный POCO и не всегда корректно уведомляет о смене вложенных свойств.
                    if (!string.IsNullOrEmpty(order.Status))
                    {
                        _viewModel.CurrentOrder.Status = order.Status;
                        StatusComboBox.SelectedValue = order.Status;
                    }

                    if (!string.IsNullOrEmpty(order.PaymentMethod))
                    {
                        _viewModel.CurrentOrder.PaymentMethod = order.PaymentMethod;
                        PaymentMethodComboBox.SelectedValue = order.PaymentMethod;
                    }

                    OrderDatePicker.SelectedDate = order.Time;
                    OrderTimeTextBox.Text = order.Time.ToString("HH:mm");

                    int branchId = order.BranchId > 0
                        ? order.BranchId
                        : (_currentShift?.BranchId ?? AppSettings.CurrentBranchId);

                    _viewModel.CurrentOrder.BranchId = branchId;
                    BranchComboBox.SelectedValue = branchId;

                    RefreshZonesForSelectedBranch(order);
                }
                else
                {
                    OrderDatePicker.SelectedDate = DateTime.Now;
                    OrderTimeTextBox.Text = DateTime.Now.ToString("HH:mm");

                    var firstWasher = allUsers?.FirstOrDefault();

                    if (firstWasher != null)
                    {
                        _viewModel.SelectedWasherId = firstWasher.Id;
                    }

                    int branchId = _viewModel.CurrentOrder.BranchId > 0
                        ? _viewModel.CurrentOrder.BranchId
                        : (_currentShift?.BranchId ?? AppSettings.CurrentBranchId);

                    _viewModel.CurrentOrder.BranchId = branchId;
                    BranchComboBox.SelectedValue = branchId;

                    RefreshZonesForSelectedBranch(null);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка загрузки данных с сервера: {ex.Message}",
                    "Ошибка API",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        // Метод генерации боксов для выбранного филиала
        private void UpdateZones(ContractsOrder order)
        {
            if (_viewModel?.CurrentOrder == null)
                return;

            if (BranchComboBox.SelectedValue is not int branchId)
                return;

            var currentBranch = _allBranches?.FirstOrDefault(b => b.Id == branchId);
            var allZones = currentBranch != null
                ? BuildZones(currentBranch)
                : new List<ZoneItem>();

            // ПЕРЕДАЕМ все зоны в ViewModel для фильтрации
            _viewModel.UpdateAvailableZones(allZones);

            // Пытаемся выбрать зону
            if (order != null &&
                _viewModel.AvailableZones.Any(z => z.BoxNumber == order.BoxNumber && z.Department == order.Department))
            {
                ZoneComboBox.SelectedValue = order.BoxNumber;
            }
            else if (_viewModel.AvailableZones.Any())
            {
                // Если заказ новый или старая зона не подходит - выбираем ПЕРВУЮ доступную
                var firstZone = _viewModel.AvailableZones.First();

                ZoneComboBox.SelectedValue = firstZone.BoxNumber;
                _viewModel.CurrentOrder.BoxNumber = firstZone.BoxNumber;
            }
            else
            {
                ZoneComboBox.ClearSelection();
            }
        }

        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Собираем время
                if (OrderDatePicker.SelectedDate.HasValue)
                {
                    if (TimeSpan.TryParse(OrderTimeTextBox.Text, out var time))
                    {
                        DateTime localTime = OrderDatePicker.SelectedDate.Value.Date + time;
                        _viewModel.CurrentOrder.Time = localTime;
                    }
                    else
                    {
                        MessageBox.Show("Неверный формат времени (ЧЧ:ММ)");
                        return;
                    }
                }

                // 2. Собираем данные из комбобоксов
                if (WasherComboBox.SelectedValue is int washerId)
                {
                    _viewModel.CurrentOrder.SetWasherId(washerId);
                }

                if (ClientComboBox.SelectedValue is int clientId)
                {
                    _viewModel.CurrentOrder.ClientId = clientId;
                }

                if (StatusComboBox.SelectedValue is string status && !string.IsNullOrWhiteSpace(status))
                {
                    _viewModel.CurrentOrder.Status = status;
                }
                else if (string.IsNullOrEmpty(_viewModel.CurrentOrder.Status))
                {
                    _viewModel.CurrentOrder.Status = "В работе";
                }

                if (PaymentMethodComboBox.SelectedValue is string paymentMethod && !string.IsNullOrWhiteSpace(paymentMethod))
                {
                    _viewModel.CurrentOrder.PaymentMethod = paymentMethod;
                }

                // 3. ЧИТАЕМ ДАННЫЕ ИЗ ВЫБРАННОЙ ЗОНЫ
                if (ZoneComboBox.SelectedValue is int selectedBoxNumber)
                {
                    var selectedZone = _viewModel.AvailableZones?
                        .FirstOrDefault(z => z.BoxNumber == selectedBoxNumber);

                    if (selectedZone != null)
                    {
                        _viewModel.CurrentOrder.BoxNumber = selectedZone.BoxNumber;
                        _viewModel.CurrentOrder.Department = selectedZone.Department;
                    }
                }

                // Подстраховка: если BranchId почему-то остался нулем, берем текущий филиал
                if (_viewModel.CurrentOrder.BranchId <= 0)
                {
                    _viewModel.CurrentOrder.BranchId = AppSettings.CurrentBranchId;
                }

                // ПРОВЕРКА ДОСТУПНОСТИ ДЛЯ ЗАПИСЕЙ
                if (_viewModel.CurrentOrder.IsAppointment)
                {
                    if (int.TryParse(DurationTextBox.Text, out int duration))
                        _viewModel.CurrentOrder.DurationMinutes = duration;
                    else
                        _viewModel.CurrentOrder.DurationMinutes = 60; // По умолчанию

                    bool isAvailable = await _apiService.CheckBoxAvailabilityForAppointmentAsync(
                        _viewModel.CurrentOrder.BranchId,       // 1. Филиал
                        _viewModel.CurrentOrder.BoxNumber,      // 2. Номер бокса
                        _viewModel.CurrentOrder.Time,           // 3. Время начала
                        _viewModel.CurrentOrder.DurationMinutes,// 4. Длительность
                        _viewModel.CurrentOrder.Id);            // 5. Игнорируем сам заказ при редактировании

                    if (!isAvailable)
                    {
                        MessageBox.Show(
                            $"Время {_viewModel.CurrentOrder.Time:HH:mm} в выбранной зоне уже занято!",
                            "Ошибка",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return; // Останавливаем сохранение
                    }
                }

                this.IsEnabled = false; // Блокируем UI на время запроса

                if (_viewModel.CurrentOrder.IsAppointment)
                {
                    _viewModel.CurrentOrder.ShiftId = 0;
                }
                else
                {
                    // Если это заказ, смена должна быть открыта
                    if (_currentShift != null && !_currentShift.IsClosed)
                    {
                        _viewModel.CurrentOrder.ShiftId = _currentShift.Id;
                        _viewModel.CurrentOrder.BranchId = _currentShift.BranchId;
                    }
                    else
                    {
                        MessageBox.Show(
                            "Нельзя сохранить заказ: на этом филиале нет активной смены!\nСначала начните смену.",
                            "Внимание",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return;
                    }
                }

                // ПРОВЕРКА: нельзя создать заказ сервиса для филиала без сервисных зон
                if (_viewModel.CurrentOrder.Department == "Service")
                {
                    var branch = _allBranches?.FirstOrDefault(b => b.Id == _viewModel.CurrentOrder.BranchId);

                    if (branch != null && branch.ServiceLiftsCount == 0)
                    {
                        MessageBox.Show(
                            $"В выбранном филиале \"{branch.Name}\" нет сервисных зон (подъемников).\n\n" +
                            $"Выберите филиал с сервисом или смените департамент на \"Мойка\".",
                            "Ошибка валидации",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);

                        return;
                    }
                }

                // 4. Вызываем асинхронное сохранение во ViewModel
                var result = await _viewModel.SaveOrderAsync();

                if (result.success)
                {
                    MessageBox.Show(result.message, "Успешно", MessageBoxButton.OK, MessageBoxImage.Information);
                    DialogResult = true;
                    Close();
                }
                else
                {
                    MessageBox.Show(result.message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Критическая ошибка: {ex.Message}");
            }
            finally
            {
                this.IsEnabled = true;
            }
        }

        private void ServicesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Как только пользователь кликает по услуге (выбирает или снимает выбор),
            // мы заставляем ViewModel пересчитать всю математику.
            if (DataContext is AddEditOrderViewModel vm)
            {
                vm.Recalculate();
            }
        }

        private void AddNewClient_Click(object sender, RoutedEventArgs e)
        {
            var win = new AddEditClientWindow(null);

            if (win.ShowDialog() == true)
            {
                // Не забываем передать CurrentOrder, как это было в твоем коде
                _ = LoadDictionariesAsync(_viewModel.CurrentOrder);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Метод для "Преобразования записи в заказ"
        private async void ConvertToOrderButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentShift == null || _currentShift.IsClosed)
            {
                MessageBox.Show(
                    "Нельзя преобразовать запись в заказ: на этом филиале нет активной смены!\nСначала начните смену.",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            var result = MessageBox.Show("Преобразовать запись в активный заказ?", "Подтверждение", MessageBoxButton.YesNo);

            if (result != MessageBoxResult.Yes)
                return;

            // Берём мойщиков прямо из кэша ViewModel, чтобы не делать лишний запрос к API
            var washers = _viewModel.Washers
                .Where(u => u.RoleId != 1 && u.RoleId != 2)
                .ToList();

            if (!washers.Any())
            {
                MessageBox.Show("Нет доступных мойщиков!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var washerDialog = new WasherSelectionDialog(washers);

            if (washerDialog.ShowDialog() != true || washerDialog.SelectedWasher == null)
                return;

            var selectedWasher = washerDialog.SelectedWasher;

            try
            {
                this.IsEnabled = false;

                if (_viewModel.CurrentOrder.Id > 0)
                {
                    // Запись УЖЕ в базе. Делаем конвертацию строго через сервер
                    await _apiService.ConvertAppointmentToOrderAsync(
                        _viewModel.CurrentOrder.Id,
                        _currentShift.Id,
                        selectedWasher.Id);

                    MessageBox.Show(
                        $"✅ Запись преобразована в заказ!\nМойщик назначен: {selectedWasher.FullName}",
                        "Успешно",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    // Закрываем окно. Заказ сохранен, UI на главной обновится по SignalR или вручную
                    DialogResult = true;
                    Close();
                }
                else
                {
                    // Это НОВАЯ запись, которой еще нет в базе (нажали Плюс -> Преобразовать).
                    // Просто меняем свойства вью-модели и вызываем обычное сохранение
                    _viewModel.SelectedWasherId = selectedWasher.Id;
                    _viewModel.CurrentOrder.IsAppointment = false;
                    _viewModel.CurrentOrder.Status = "В работе";
                    _viewModel.SetAsOrder();

                    SaveButton_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при преобразовании: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                this.IsEnabled = true;
            }
        }

        // Обработчик поиска по услугам
        private void ServiceSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DataContext is AddEditOrderViewModel vm)
            {
                vm.ServiceSearchText = ServiceSearchTextBox.Text;
            }
        }

        // === ОБРАБОТЧИК ДОБАВЛЕНИЯ РАСХОДА ===
        private async void AddExpenseButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ExpenseNameTextBox.Text))
            {
                MessageBox.Show("Введите название расхода", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!decimal.TryParse(ExpenseCostPriceTextBox.Text.Replace(",", "."), out decimal costPrice) || costPrice < 0)
            {
                MessageBox.Show("Введите корректную себестоимость", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!decimal.TryParse(ExpenseClientPriceTextBox.Text.Replace(",", "."), out decimal clientPrice) || clientPrice < 0)
            {
                MessageBox.Show("Введите корректную цену для клиента", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(ExpenseQuantityTextBox.Text, out int quantity) || quantity <= 0)
            {
                quantity = 1;
            }

            var categoryItem = ExpenseCategoryComboBox.SelectedItem as ComboBoxItem;

            if (categoryItem == null)
            {
                MessageBox.Show("Выберите категорию", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string category = ExpenseCategoryComboBox.SelectedValue as string;

            if (string.IsNullOrWhiteSpace(category))
            {
                MessageBox.Show(
                    "Выберите категорию",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            string note = string.IsNullOrWhiteSpace(ExpenseNoteTextBox.Text)
                ? ""
                : ExpenseNoteTextBox.Text;

            this.IsEnabled = false;

            bool success = await _viewModel.AddExpenseAsync(
                ExpenseNameTextBox.Text,
                category,
                costPrice,
                clientPrice,
                quantity,
                note);

            this.IsEnabled = true;

            if (success)
            {
                MessageBox.Show("Расход добавлен!", "Успешно", MessageBoxButton.OK, MessageBoxImage.Information);

                // Очищаем поля
                ExpenseNameTextBox.Clear();
                ExpenseCostPriceTextBox.Clear();
                ExpenseClientPriceTextBox.Clear();
                ExpenseQuantityTextBox.Text = "1";
                ExpenseNoteTextBox.Clear();

                // ИСПРАВЛЕНО: SelectedItem нельзя выставить в -1.
                // Сброс выбора делается через SelectedIndex.
                ExpenseCategoryComboBox.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Не удалось добавить расход", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PrintActButton_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.CurrentOrder == null || _viewModel.CurrentOrder.Id <= 0)
            {
                MessageBox.Show("Сохраните заказ перед печатью акта", "Предупреждение");
                return;
            }

            MessageBox.Show(
                "Функция печати акта пока не реализована на клиенте.\n\nДанные для акта:\n" +
                $"Заказ #{_viewModel.CurrentOrder.Id}\n" +
                $"Клиент: {_viewModel.CurrentOrder.ClientId}\n" +
                $"Услуги: {_viewModel.Services?.Count(s => s.IsSelected)} шт.\n" +
                $"Расходы: {_viewModel.OrderExpenses?.Count} шт.\n" +
                $"Итого: {_viewModel.FinalTotalWithExpenses:N0} ₽",
                "Печать акта",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ApplyUpsell_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is AddEditOrderViewModel vm)
            {
                vm.ApplyUpsell();
            }
        }
    }

    public class ZoneItem
    {
        public string Name { get; set; }
        public int BoxNumber { get; set; }
        public string Department { get; set; }
    }
}