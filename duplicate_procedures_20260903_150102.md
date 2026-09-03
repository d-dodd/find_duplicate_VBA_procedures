# Duplicate VBA Procedures

Source folder: `C:\Work in Progress\Americas Model\GreenArrow\GA testing & development\Core Models\VBA_Export`

Generated: 2026-09-03 15:01:02

- Files scanned: 142
- Live procedure declarations: 2092
- Commented-out declarations ignored: 376
- Duplicates in the same module: 0
- Ambiguous public names across standard modules: 8
- Repeated names across modules (legal): 18

A name can appear in more than one section below: a routine declared twice in one module and again in a second module is both a compile error and an ambiguity.

## Duplicates in the same module

None found.

## Ambiguous public names across standard modules

The same public name in more than one standard (.bas) module. VBA raises "Ambiguous name detected" unless every call site is qualified with the module name. Make all but one Private, or rename.

### `Copy_KeyAssumptions_AM` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| DataTapeDelivery | Standard module | DataTapeDelivery.bas | 232 | Sub | Public (Option Private Module) | `Sub Copy_KeyAssumptions_AM()` |
| KeyAssumptionsMacros | Standard module | KeyAssumptionsMacros.bas | 89 | Sub | Public (Option Private Module) | `Sub Copy_KeyAssumptions_AM()` |

### `fnAlreadyOpen` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| Self_Storage_Macros | Standard module | Self_Storage_Macros.bas | 775 | Function | Public (Option Private Module) | `Function fnAlreadyOpen(Fname As String) As Boolean` |
| Shared_Functions | Standard module | Shared_Functions.bas | 821 | Function | Public (Option Private Module) | `Function fnAlreadyOpen(Fname As String) As Boolean` |

### `fnArrayIndx` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| Self_Storage_Macros | Standard module | Self_Storage_Macros.bas | 525 | Function | Public (Option Private Module) | `Function fnArrayIndx(min As Long, max As Long, Optional pfx As String) As Variant` |
| Shared_Functions | Standard module | Shared_Functions.bas | 153 | Function | Public (Option Private Module) | `Function fnArrayIndx(min As Long, max As Long) As Variant` |

### `fnOnlyBefore` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| Self_Storage_Macros | Standard module | Self_Storage_Macros.bas | 546 | Function | Public (Option Private Module) | `Function fnOnlyBefore(str As String, substr As String) As String` |
| Shared_Functions | Standard module | Shared_Functions.bas | 1695 | Function | Public (Option Private Module) | `Function fnOnlyBefore(str As String, substr As String) As String` |

### `LoadFile` (3 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| L_sheet_macros | Standard module | L_sheet_macros.bas | 49 | Sub | Public | `Sub LoadFile(FileName1 As String, Filename2 As String)` |
| Macros | Standard module | Macros.bas | 9600 | Sub | Public (Option Private Module) | `Sub LoadFile(FileName1 As String, Filename2 As String)` |
| rent_routines_functions | Standard module | rent_routines_functions.bas | 279 | Sub | Public | `Sub LoadFile(FileName1 As String, Filename2 As String)` |

### `NameExists` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| Exp_Charts | Standard module | Exp_Charts.bas | 261 | Function | Public (Option Private Module) | `Function NameExists(nN As String, Optional sc As String) As Boolean` |
| Self_Storage_Macros | Standard module | Self_Storage_Macros.bas | 388 | Function | Public (Option Private Module) | `Function NameExists(myName As Variant) As Boolean` |

### `Send_KeyAssumptions_AM` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| DataTapeDelivery | Standard module | DataTapeDelivery.bas | 4 | Sub | Public (Option Private Module) | `Sub Send_KeyAssumptions_AM()` |
| KeyAssumptionsMacros | Standard module | KeyAssumptionsMacros.bas | 4 | Sub | Public (Option Private Module) | `Sub Send_KeyAssumptions_AM()` |

### `toggleTypeSubtypeRent` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| changeRangeValues | Standard module | changeRangeValues.bas | 150 | Sub | Public (Option Private Module) | `Sub toggleTypeSubtypeRent()` |
| rent_routines_functions | Standard module | rent_routines_functions.bas | 41 | Sub | Public | `Sub toggleTypeSubtypeRent()` |

## Repeated names across modules

Legal in VBA because the names are scoped to a class, form, or to a single module via Private/Friend. Worth reviewing anyway: these are usually copy-pasted routines that have since drifted apart.

### `btnExit_Click` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| frmMarket | Form module | frmMarket.frm | 270 | Sub | Private | `Private Sub btnExit_Click()` |
| frmOpExp | Form module | frmOpExp.frm | 119 | Sub | Private | `Private Sub btnExit_Click()` |

### `cancelButton_Click` (3 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| IS_Grid_format | Form module | IS_Grid_format.frm | 20 | Sub | Private | `Private Sub cancelButton_Click()` |
| numberBuildingComponents | Form module | numberBuildingComponents.frm | 19 | Sub | Private | `Private Sub cancelButton_Click()` |
| numberBuildingComponentsCost | Form module | numberBuildingComponentsCost.frm | 19 | Sub | Private | `Private Sub cancelButton_Click()` |

### `Center_MainForm` (3 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| frmMain | Form module | frmMain.frm | 457 | Sub | Private | `Private Sub Center_MainForm()` |
| frmMarket | Form module | frmMarket.frm | 281 | Sub | Private | `Private Sub Center_MainForm()` |
| NISImportConfig | Form module | NISImportConfig.frm | 30 | Sub | Private | `Private Sub Center_MainForm()` |

### `center_size` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| expenseAnalysisUnitOfComparison | Form module | expenseAnalysisUnitOfComparison.frm | 36 | Sub | Private | `Private Sub center_size(h As Double, w As Double)` |
| frmUnitMix | Form module | frmUnitMix.frm | 42 | Sub | Private | `Private Sub center_size(h As Double, w As Double)` |

### `Class_Initialize` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| AnalyticsApp | Class module | AnalyticsApp.cls | 19 | Sub | Private | `Private Sub Class_Initialize()` |
| IntroStateManager | Class module | IntroStateManager.cls | 30 | Sub | Private | `Private Sub Class_Initialize()` |

### `Class_Terminate` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| AnalyticsApp | Class module | AnalyticsApp.cls | 24 | Sub | Private | `Private Sub Class_Terminate()` |
| IntroStateManager | Class module | IntroStateManager.cls | 47 | Sub | Private | `Private Sub Class_Terminate()` |

### `CommandButton1_Click` (6 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| expenseAnalysisUnitOfComparison | Form module | expenseAnalysisUnitOfComparison.frm | 18 | Sub | Private | `Private Sub CommandButton1_Click()` |
| Market_Rent | Form module | Market_Rent.frm | 35 | Sub | Private | `Private Sub CommandButton1_Click()` |
| numberCostComps | Form module | numberCostComps.frm | 19 | Sub | Private | `Private Sub CommandButton1_Click()` |
| OpExpenseConfig | Form module | OpExpenseConfig.frm | 67 | Sub | Private | `Private Sub CommandButton1_Click()` |
| RentSumDisplay | Form module | RentSumDisplay.frm | 18 | Sub | Private | `Private Sub CommandButton1_Click()` |
| ttChangeDataForm | Form module | ttChangeDataForm.frm | 20 | Sub | Private | `Private Sub CommandButton1_Click()` |

### `CommandButton2_Click` (5 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| expenseAnalysisUnitOfComparison | Form module | expenseAnalysisUnitOfComparison.frm | 32 | Sub | Private | `Private Sub CommandButton2_Click()` |
| Market_Rent | Form module | Market_Rent.frm | 64 | Sub | Private | `Private Sub CommandButton2_Click()` |
| numberCostComps | Form module | numberCostComps.frm | 30 | Sub | Private | `Private Sub CommandButton2_Click()` |
| OpExpenseConfig | Form module | OpExpenseConfig.frm | 19 | Sub | Private | `Private Sub CommandButton2_Click()` |
| ttChangeDataForm | Form module | ttChangeDataForm.frm | 35 | Sub | Private | `Private Sub CommandButton2_Click()` |

### `ConnectionString` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| clsMktModelsDB | Class module | clsMktModelsDB.cls | 34 | Function | Public | `Public Function ConnectionString()` |
| clsModelsDB | Class module | clsModelsDB.cls | 33 | Function | Public | `Public Function ConnectionString()` |

### `DatabaseName` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| clsMktModelsDB | Class module | clsMktModelsDB.cls | 30 | Function | Private | `Private Function DatabaseName()` |
| clsModelsDB | Class module | clsModelsDB.cls | 29 | Function | Private | `Private Function DatabaseName()` |

### `FillComboBox` (3 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| frmMain | Form module | frmMain.frm | 244 | Sub | Private | `Private Sub FillComboBox(cbx As ComboBox, query As String, col As String, conn As Object)` |
| frmMarket | Form module | frmMarket.frm | 615 | Sub | Private | `Private Sub FillComboBox(cbx As ComboBox, query As String, col As String, conn As Object)` |
| NISImportConfig | Form module | NISImportConfig.frm | 159 | Sub | Private | `Private Sub FillComboBox(cbx As ComboBox, query As String, col As String, conn As Object)` |

### `okButton_Click` (4 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| IS_Grid_format | Form module | IS_Grid_format.frm | 23 | Sub | Private | `Private Sub okButton_Click()` |
| numberBuildingComponents | Form module | numberBuildingComponents.frm | 23 | Sub | Private | `Private Sub okButton_Click()` |
| numberBuildingComponentsCost | Form module | numberBuildingComponentsCost.frm | 23 | Sub | Private | `Private Sub okButton_Click()` |
| RentSumDisplay | Form module | RentSumDisplay.frm | 22 | Sub | Private | `Private Sub okButton_Click()` |

### `Server` (2 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| clsMktModelsDB | Class module | clsMktModelsDB.cls | 26 | Function | Private | `Private Function Server()` |
| clsModelsDB | Class module | clsModelsDB.cls | 25 | Function | Private | `Private Function Server()` |

### `UserForm_Activate` (13 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| CompSelectForm | Form module | CompSelectForm.frm | 45 | Sub | Private | `Private Sub UserForm_Activate()` |
| expenseAnalysisUnitOfComparison | Form module | expenseAnalysisUnitOfComparison.frm | 65 | Sub | Private | `Private Sub UserForm_Activate()` |
| frmMain | Form module | frmMain.frm | 19 | Sub | Private | `Private Sub UserForm_Activate()` |
| frmMarket | Form module | frmMarket.frm | 278 | Sub | Private | `Private Sub UserForm_Activate()` |
| IS_Grid_format | Form module | IS_Grid_format.frm | 51 | Sub | Private | `Private Sub UserForm_Activate()` |
| Market_Rent | Form module | Market_Rent.frm | 17 | Sub | Private | `Private Sub UserForm_Activate()` |
| NISImportConfig | Form module | NISImportConfig.frm | 27 | Sub | Private | `Private Sub UserForm_Activate()` |
| numberBuildingComponents | Form module | numberBuildingComponents.frm | 37 | Sub | Private | `Private Sub UserForm_Activate()` |
| numberBuildingComponentsCost | Form module | numberBuildingComponentsCost.frm | 62 | Sub | Private | `Private Sub UserForm_Activate()` |
| numberCostComps | Form module | numberCostComps.frm | 47 | Sub | Private | `Private Sub UserForm_Activate()` |
| OpExpenseConfig | Form module | OpExpenseConfig.frm | 51 | Sub | Private | `Private Sub UserForm_Activate()` |
| RentSumDisplay | Form module | RentSumDisplay.frm | 50 | Sub | Private | `Private Sub UserForm_Activate()` |
| TaxAssumptions | Form module | TaxAssumptions.frm | 57 | Sub | Private | `Private Sub UserForm_Activate()` |

### `userform_initialize` (16 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| CompSelectForm | Form module | CompSelectForm.frm | 38 | Sub | Private | `Private Sub userform_initialize()` |
| expenseAnalysisUnitOfComparison | Form module | expenseAnalysisUnitOfComparison.frm | 51 | Sub | Private | `Private Sub userform_initialize()` |
| frmCMR_Update | Form module | frmCMR_Update.frm | 82 | Sub | Private | `Private Sub userform_initialize()` |
| frmMain | Form module | frmMain.frm | 22 | Sub | Private | `Private Sub userform_initialize()` |
| frmMarket | Form module | frmMarket.frm | 296 | Sub | Private | `Private Sub userform_initialize()` |
| frmOpExp | Form module | frmOpExp.frm | 23 | Sub | Private | `Private Sub userform_initialize()` |
| frmUnitMix | Form module | frmUnitMix.frm | 39 | Sub | Public | `Sub userform_initialize()` |
| IS_Grid_format | Form module | IS_Grid_format.frm | 36 | Sub | Private | `Private Sub userform_initialize()` |
| NISImportConfig | Form module | NISImportConfig.frm | 176 | Sub | Private | `Private Sub userform_initialize()` |
| numberBuildingComponents | Form module | numberBuildingComponents.frm | 29 | Sub | Private | `Private Sub userform_initialize()` |
| numberBuildingComponentsCost | Form module | numberBuildingComponentsCost.frm | 41 | Sub | Private | `Private Sub userform_initialize()` |
| numberCostComps | Form module | numberCostComps.frm | 34 | Sub | Private | `Private Sub userform_initialize()` |
| OpExpenseConfig | Form module | OpExpenseConfig.frm | 23 | Sub | Private | `Private Sub userform_initialize()` |
| RentSumDisplay | Form module | RentSumDisplay.frm | 37 | Sub | Private | `Private Sub userform_initialize()` |
| TaxAssumptions | Form module | TaxAssumptions.frm | 43 | Sub | Private | `Private Sub userform_initialize()` |
| ttChangeDataForm | Form module | ttChangeDataForm.frm | 41 | Sub | Private | `Private Sub userform_initialize()` |

### `Worksheet_Activate` (4 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| S_Income_DirectCap_1 | Class module | S_Income_DirectCap_1.cls | 41 | Sub | Private | `Private Sub Worksheet_Activate()` |
| S_Income_DirectCap_5 | Class module | S_Income_DirectCap_5.cls | 37 | Sub | Private | `Private Sub Worksheet_Activate()` |
| S_Income_DirectCap_6 | Class module | S_Income_DirectCap_6.cls | 37 | Sub | Private | `Private Sub Worksheet_Activate()` |
| S_Income_DirectCap_7 | Class module | S_Income_DirectCap_7.cls | 37 | Sub | Private | `Private Sub Worksheet_Activate()` |

### `Worksheet_BeforeDoubleClick` (7 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| S_Cost | Class module | S_Cost.cls | 10 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_ExecutiveSummary | Class module | S_ExecutiveSummary.cls | 57 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_Improv | Class module | S_Improv.cls | 98 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_Inputs | Class module | S_Inputs.cls | 19 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_Site | Class module | S_Site.cls | 10 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_Tax | Class module | S_Tax.cls | 10 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |
| S_Zoning | Class module | S_Zoning.cls | 11 | Sub | Private | `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)` |

### `Worksheet_Change` (21 declarations)

| Module | Module type | File | Line | Kind | Scope | Declaration |
| --- | --- | --- | --- | --- | --- | --- |
| F_CatForm | Class module | F_CatForm.cls | 11 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Area | Class module | S_Area.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Construction | Class module | S_Construction.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_ExecutiveSummary | Class module | S_ExecutiveSummary.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Exp_Analysis | Class module | S_Exp_Analysis.cls | 12 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_ExportControls | Class module | S_ExportControls.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Improv | Class module | S_Improv.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Income_DirectCap_1 | Class module | S_Income_DirectCap_1.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Income_DirectCap_2 | Class module | S_Income_DirectCap_2.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Income_DirectCap_5 | Class module | S_Income_DirectCap_5.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Income_DirectCap_6 | Class module | S_Income_DirectCap_6.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Income_DirectCap_7 | Class module | S_Income_DirectCap_7.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Inputs | Class module | S_Inputs.cls | 24 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_IntroNarr | Class module | S_IntroNarr.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_LUD | Class module | S_LUD.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_MarketAnalysis | Class module | S_MarketAnalysis.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_OpHistory | Class module | S_OpHistory.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_REIS | Class module | S_REIS.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_Rent_Comparison_2 | Class module | S_Rent_Comparison_2.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_RentByType | Class module | S_RentByType.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |
| S_SubDetail | Class module | S_SubDetail.cls | 10 | Sub | Private | `Private Sub Worksheet_Change(ByVal Target As Range)` |

