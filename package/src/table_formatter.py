import unohelper
from scriptforge import CreateScriptService
from com.sun.star.awt import ImageScaleMode, Key
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import XActionListener, XItemListener, XKeyListener
import random as rnd
import json
import uno
import re

# Used to call Basic functions from within Python
bas = None
exc = None
fs = None
platform = None

# Base path to the extension root (in URL format)
base_path_url = ""
styles_path_url = ""
previews_path_url = ""
config_file_url = ""
styles_file_url = ""
template_odp_url = ""
po_folder_url = ""
temp_folder_url = ""

# Base path to the extension root (in SYS format)
base_path_sys = ""
styles_path_sys = ""
previews_path_sys = ""
config_file_sys = ""
styles_file_sys = ""
template_odp_sys = ""
po_folder_sys = ""
temp_folder_url = ""

# Store globally the most recent json and png created (in URL format)
current_json_temp_url = ""
current_png_temp_url = ""

# After closing the New Style dialog, indicates if the list box needs reloading
b_needs_reloading = False
s_new_style_name = ""

# Indicates whether L10n and Paths need to be initialized
b_needs_initialize_l10n = True

# Stores the L10N service instance used for localization
l10n = None

# Stores the current locale of the l10n service
l10n_locale = ""

# Indicates whether style names need to be checked for translation
need_style_translation = True

# Global variable that stores all the available styles
G_STYLES = None

# Global variable that indicates that the config file needs to be saved
G_SAVE_CONFIG = False

# Stores the dialog options in the current session
G_DLG_OPTIONS = None

# Defines the maximum number of custom styles allowed
CUSTOM_STYLE_LIMIT = 1000


# Initialize all paths for the extension
def initialize_paths():
    # Saves current FileNaming property to not mess with any other running macro
    previous_filenaming = fs.FileNaming
    fs.FileNaming = "URL"
    # Initialize base path to the extension root (in URL format)
    global base_path_url, styles_path_url, previews_path_url, config_file_url
    global styles_file_url, template_odp_url, po_folder_url, temp_folder_url
    # This is the only path that is initialized; all other are relative to base_path_url
    # Using GetDefaultContext instead of ExtensionFolder for compatibility with 7.3.x
    # base_path_url = fs.ExtensionFolder("org.rafaelhlima.ImpressTableFormatter")
    pip = bas.GetDefaultContext().getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    base_path_url = pip.getPackageLocation("org.rafaelhlima.ImpressTableFormatter")
    config_file_url = fs.BuildPath(base_path_url, "all_styles.json")
    previews_path_url = fs.BuildPath(base_path_url, "preview_files")
    styles_file_url = fs.BuildPath(base_path_url, "all_styles.json")
    styles_path_url = fs.BuildPath(base_path_url, "format_files")
    template_odp_url = fs.BuildPath(base_path_url, "temp")
    template_odp_url = fs.BuildPath(template_odp_url, "template_for_png.odp")
    po_folder_url = fs.BuildPath(base_path_url, "po")
    temp_folder_url = fs.BuildPath(base_path_url, "temp")
    # Base path to the extension root (in SYS format)
    global base_path_sys, styles_path_sys, previews_path_sys, config_file_sys
    global styles_file_sys, template_odp_sys, po_folder_sys, temp_folder_sys
    base_path_sys = bas.ConvertFromUrl(base_path_url)
    config_file_sys = bas.ConvertFromUrl(config_file_url)
    previews_path_sys = bas.ConvertFromUrl(previews_path_url)
    styles_file_sys = bas.ConvertFromUrl(styles_file_url)
    styles_path_sys = bas.ConvertFromUrl(styles_path_url)
    template_odp_sys = bas.ConvertFromUrl(template_odp_url)
    po_folder_sys = bas.ConvertFromUrl(po_folder_url)
    temp_folder_sys = bas.ConvertFromUrl(temp_folder_url)
    # Restore the FileNaming property (don't mess with anything else running
    # that relies on FileSystem service)
    fs.FileNaming = previous_filenaming


# Initialize the L10N service for localization
def initialize_l10n():
    global l10n
    try:
        # Available since LO 7.4 (uses LO User Interface langage)
        current_locale = platform.OfficeLocale
    except Exception:
        # Gets the operating system locale
        current_locale = platform.Locale
    # Check if the file for the current OfficeLocale exists
    po_file = fs.BuildPath(po_folder_url, current_locale + ".po")
    # Generic language file la.po (may be necessary if la-.po does not exist)
    language = current_locale.split("-")[0]
    lang_file = language + ".po"
    lang_file_path = fs.BuildPath(po_folder_url, lang_file)
    if fs.FileExists(po_file):
        # If the po file exists, then use it
        l10n = CreateScriptService("L10N", po_folder_url)
    elif fs.FileExists(lang_file_path):
        # Check if there is a generic language la.po file
        l10n = CreateScriptService("L10N", po_folder_url, language)
    else:
        # If it does not exist, uses en-US.po (default language, always exist)
        l10n = CreateScriptService("L10N", po_folder_url, "en-US")


# Natural sort used to sort style names
def natural_sort(l):
    convert = lambda text: int(text) if text.isdigit() else text.lower()
    alphanum_key = lambda key: [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(l, key=alphanum_key)


# Return a tuple (r, g, b) from an integer color
def get_rgb_from_color(color):
    red = color >> 16
    color_updated = color - red * 65536
    green = color_updated >> 8
    blue = color_updated - green * 256
    return (red, green, blue)


# Convert a integer color to a dictionary
def color_to_dict(color):
    components = get_rgb_from_color(color)
    rgb_dict = dict()
    rgb_dict["red"] = components[0]
    rgb_dict["green"] = components[1]
    rgb_dict["blue"] = components[2]
    return rgb_dict


# Return (r, g, b) values from RGB dict
def rgb_from_dict(color_dict):
    return color_dict["red"], color_dict["green"], color_dict["blue"]


# Get the format of a single border
def get_border_format(border):
    format = dict()
    format["line-color"] = color_to_dict(border.Color)
    format["line-style"] = border.LineStyle
    format["line-width"] = border.LineWidth
    return format


# Get format from single cell
def get_cell_format(cell):
    cell_format = dict()
    # Cell background color
    cell_format["bg-color"] = color_to_dict(cell.FillColor)
    cell_format["bg-color2"] = color_to_dict(cell.FillColor2)
    # Cell background style
    cell_format["fill-style"] = cell.FillStyle.value
    cell_format["fill-transparence"] = cell.FillTransparence
    # Font information
    cell_format["font-name"] = cell.CharFontName
    cell_format["font-size"] = cell.CharHeight
    cell_format["font-weight"] = cell.CharWeight
    # Font color (test if color is automatic)
    if cell.CharColor != -1:
        # There is a RGB color defined
        cell_format["font-color"] = color_to_dict(cell.CharColor)
    else:
        # Color = Automatic
        cell_format["font-color"] = -1
    # Text alignment
    cell_format["text-h-align"] = cell.TextHorizontalAdjust.value
    cell_format["text-v-align"] = cell.TextVerticalAdjust.value
    cell_format["para-adjust"] = cell.ParaAdjust
    # Border format
    cell_format["bottom-border"] = get_border_format(cell.BottomBorder)
    cell_format["left-border"] = get_border_format(cell.LeftBorder)
    cell_format["right-border"] = get_border_format(cell.RightBorder)
    # cell_format["top-border"] = get_border_format(cell.BottomBorder)
    return cell_format


# Return the format of borders around the table. Considers the top-left,
# and bottom-right borders as reference to get the formats
def get_table_borders(table):
    nCols = table.ColumnCount
    nRows = table.RowCount
    table_borders = dict()
    # Top-border of table = top border of top-left cell
    # Left-border of table = left border of top-left cell
    cell = table.getCellByPosition(0, 0)
    table_borders["top-border"] = get_border_format(cell.TopBorder)
    table_borders["left-border"] = get_border_format(cell.LeftBorder)
    # Bottom-border of table = bottom border of bottom-right cell
    # Right-border of table = right border of bottom-right cell
    cell = table.getCellByPosition(nCols - 1, nRows - 1)
    table_borders["bottom-border"] = get_border_format(cell.BottomBorder)
    table_borders["right-border"] = get_border_format(cell.RightBorder)
    # Return the format
    return table_borders


# Create a Json file containing all the format of a table
# The table is assumed to have at least 3 columns and 3 rows
# The second column is used to get the cell-wise formats
def create_json_from_table(table, json_file):
    table_format = dict()
    # Get the header format
    table_format["header-row"] = get_cell_format(table.getCellByPosition(1, 0))
    # Get the format for even rows
    table_format["banded-rows"] = get_cell_format(table.getCellByPosition(1, 1))
    # Get the format for odd rows
    table_format["normal-rows"] = get_cell_format(table.getCellByPosition(1, 2))
    # Get the border format to be drawn around the table
    table_format["table-borders"] = get_table_borders(table)
    # Write format to Json file
    with open(json_file, "w+", encoding="utf8") as f:
        json.dump(table_format, f, indent=2)


# Create a BorderLine2 struct with the values for one of the border lines
def create_border_line_style(border_format):
    line_format = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    r, g, b = rgb_from_dict(border_format["line-color"])
    line_format.Color = bas.RGB(r, g, b)
    line_format.LineStyle = border_format["line-style"]
    line_format.LineWidth = border_format["line-width"]
    return line_format


# Apply a row format to a single rows
def apply_row_format(table, row_id, row_format):
    nCols = table.ColumnCount
    # Format each cell in the row
    for col_id in range(nCols):
        cell = table.getCellByPosition(col_id, row_id)
        # Cell background color
        r, g, b = rgb_from_dict(row_format["bg-color"])
        cell.FillColor = bas.RGB(r, g, b)
        r, g, b = rgb_from_dict(row_format["bg-color2"])
        cell.FillColor2 = bas.RGB(r, g, b)
        # Cell background style
        bg_style = uno.Enum("com.sun.star.drawing.FillStyle", row_format["fill-style"])
        cell.FillStyle = bg_style
        cell.FillTransparence = row_format["fill-transparence"]
        # Adjustment information
        h_align = uno.Enum("com.sun.star.drawing.TextHorizontalAdjust", row_format["text-h-align"])
        cell.TextHorizontalAdjust = h_align
        v_align = uno.Enum("com.sun.star.drawing.TextVerticalAdjust", row_format["text-v-align"])
        cell.TextVerticalAdjust = v_align
        cell.ParaAdjust = row_format["para-adjust"]
        # Font name, size and weight
        cell.CharFontName = row_format["font-name"]
        cell.CharHeight = row_format["font-size"]
        cell.CharWeight = row_format["font-weight"]
        # Font color
        if row_format["font-color"] == -1:
            cell.CharColor = -1
        else:
            r, g, b = rgb_from_dict(row_format["font-color"])
            cell.CharColor = bas.RGB(r, g, b)
        # Border format
        cell.BottomBorder = create_border_line_style(row_format["bottom-border"])
        cell.LeftBorder = create_border_line_style(row_format["left-border"])
        cell.RightBorder = create_border_line_style(row_format["right-border"])
        # cell.TopBorder = create_border_line_style(row_format["top-border"])


# Clear all borders in the table
def clear_all_table_borders(table):
    # Create a new TableBorder object with default borders (all with zero)
    new_table_border = uno.createUnoStruct("com.sun.star.table.TableBorder")
    new_table_border.IsTopLineValid = True
    new_table_border.IsBottomLineValid = True
    new_table_border.IsLeftLineValid = True
    new_table_border.IsRightLineValid = True
    new_table_border.IsHorizontalLineValid = True
    new_table_border.IsVerticalLineValid = True
    new_table_border.IsDistanceValid = True
    # Number of rows and columns in the table
    nRows = table.RowCount
    nCols = table.ColumnCount
    # Apply the empty TableBorder to all cells
    for i in range(nCols):
        for j in range(nRows):
            cell = table.getCellByPosition(i, j)
            cell.TableBorder = new_table_border


# Apply a format stored in a Json file to a table
# The path to the format file must be expressed in SYS format
def apply_format_to_table(table, json_file, options):
    # Read the format from the Json file
    with open(json_file, "r", encoding="utf8") as f:
        table_format = json.load(f)
    # Remove all borders from the current table
    clear_all_table_borders(table)
    # Number of rows and columns in the table
    nRows = table.RowCount
    nCols = table.ColumnCount
    # Format the first row (may be header or not)
    if options["opt_header"] == 1:
        apply_row_format(table, 0, table_format["header-row"])
    else:
        apply_row_format(table, 0, table_format["normal-rows"])
    # Format remaining rows (check if banded rows are desired)
    if options["opt_banded"] == 1:
        # Banded rows
        for row_id in range(1, nRows):
            if row_id % 2 == 1:
                apply_row_format(table, row_id, table_format["banded-rows"])
            else:
                apply_row_format(table, row_id, table_format["normal-rows"])
    else:
        # There are no banded rows
        for row_id in range(1, nRows):
            apply_row_format(table, row_id, table_format["normal-rows"])
    # Format table outer borders
    bottom_border = create_border_line_style(table_format["table-borders"]["bottom-border"])
    left_border = create_border_line_style(table_format["table-borders"]["left-border"])
    right_border = create_border_line_style(table_format["table-borders"]["right-border"])
    top_border = create_border_line_style(table_format["table-borders"]["top-border"])
    for i in range(nRows):
        cell = table.getCellByPosition(0, i)
        cell.LeftBorder = left_border
        cell = table.getCellByPosition(nCols - 1, i)
        cell.RightBorder = right_border
    for i in range(nCols):
        cell = table.getCellByPosition(i, 0)
        cell.TopBorder = top_border
        cell = table.getCellByPosition(i, nRows - 1)
        cell.BottomBorder = bottom_border


# Minimize the height of all rows. It requires the XShape object containing the table
def compact_table_height(table_shape):
    table = table_shape.Model
    nRows = table.Rows.getCount()
    nCols = table.Columns.getCount()
    total_height = 0
    # Change the size of all rows
    for i in range(nRows):
        cell = table.getCellByPosition(0, i)
        cell_sizes = list()
        for j in range(nCols):
            cell_sizes.append(cell.MinimumSize.Height)
        min_height = max(cell_sizes)
        table.Rows.getByIndex(i).Height = min_height
        total_height += min_height
    # Resize the shape object
    newSize = uno.createUnoStruct("com.sun.star.awt.Size")
    newSize.Height = total_height
    newSize.Width = table_shape.Size.Width
    table_shape.Size = newSize


# Applies a slight gray border around the table (harcoded format)
def apply_table_border(table_shape):
    table_shape.Shadow = True
    table_shape.ShadowBlur = 176
    table_shape.ShadowColor = bas.RGB(102, 102, 102)
    table_shape.ShadowXDistance = 0
    table_shape.ShadowYDistance = 0


# Check if the current selection contains a single table
def validate_selection(doc):
    # Check if document is Impress or Draw
    ui = CreateScriptService("UI", doc)
    # bas.MsgBox(ui.ActiveWindow)
    s_doc = CreateScriptService("Document", ui.ActiveWindow)
    doc_type = s_doc.DocumentType
    # bas.MsgBox(doc_type)
    if not(doc_type == "Impress" or doc_type == "Draw"):
        msg = l10n._("Msg.Invalid_A")
        return False, msg
    currentSel = doc.getCurrentSelection()
    if currentSel is None:
        msg = l10n._("Msg.Invalid_B")
        return False, msg
    # If multiple objects are selected, nothing can be done
    if currentSel.getCount() > 1:
        msg = l10n._("Msg.Invalid_C")
        return False, msg
    # Get the model of the first selected object
    table_model = currentSel.getByIndex(0).Model
    # Check if the selected object is a table
    table_idl = "com.sun.star.table.XTable"
    is_table = False
    for interface in table_model.Types:
        if interface.typeName == table_idl:
            is_table = True
    if not is_table:
        msg = l10n._("Msg.Invalid_B")
        return False, msg
    # It it gets here the a single table is selected
    return True, "OK"


# Called when any of the checkboxes is clicked
def on_checkbox_click(event=None):
    chk_control = CreateScriptService("DialogEvent", event)
    dialog = chk_control.Parent
    # Save the current options in the checkboxes
    global G_DLG_OPTIONS
    G_DLG_OPTIONS = get_dlg_options_state(dialog)


# Creates a description string to be saved in the Description field of the table
def create_description_string(style_name, dlg_options):
    description_string = style_name
    description_string += "|" + str(G_DLG_OPTIONS["opt_header"])
    description_string += "|" + str(G_DLG_OPTIONS["opt_banded"])
    description_string += "|" + str(G_DLG_OPTIONS["opt_font"])
    description_string += "|" + str(G_DLG_OPTIONS["opt_text_align"])
    description_string += "|" + str(G_DLG_OPTIONS["opt_compact"])
    description_string += "|" + str(G_DLG_OPTIONS["opt_shadow"])
    return description_string


# Returns the name based on a Description string
def get_name_from_string(description_string):
    return description_string.strip().split("|")[0]


# Procedure called when the Apply button is pressed
def cmd_apply_pressed(event=None):
    # Gets the current document
    # doc = XSCRIPTCONTEXT.getDocument()
    doc = bas.ThisComponent
    # Gets the selected table Shape and Model
    table_shape = doc.getCurrentSelection().getByIndex(0)
    table_model = table_shape.Model
    # Gets the selected style
    btn_apply = CreateScriptService("DialogEvent", event)
    main_dialog = btn_apply.Parent
    styles_list = main_dialog.Controls("StyleList")
    selected_index = styles_list.ListIndex
    selected_style = styles_list.XControlModel.getItemData(selected_index)
    global styles_path_url
    style_file_url = fs.BuildPath(styles_path_url, G_STYLES[selected_style]["json-file"])
    style_file_sys = bas.ConvertFromUrl(style_file_url)
    # Apply formats to the table based on the dialog options
    global G_DLG_OPTIONS
    apply_format_to_table(table_model, style_file_sys, G_DLG_OPTIONS)
    # Save the applied style to the "Description" field of the table Shape
    table_shape.Description = create_description_string(selected_style, G_DLG_OPTIONS)
    # Compact the table height if required
    if G_DLG_OPTIONS["opt_compact"] == 1:
        compact_table_height(table_shape)
    # Apply a slight table shadow if required
    if G_DLG_OPTIONS["opt_shadow"] == 1:
        apply_table_border(table_shape)
    else:
        table_shape.Shadow = False
    # Close the dialog if the "Close after applying" option is checked
    if G_DLG_OPTIONS["opt_close"] == 1:
        # Using XDialogView for compatibility with LO 7.3
        # main_dialog.endExecute(0)
        main_dialog.XDialogView.endExecute()
        main_dialog.Terminate()


# Save the all_styles.json file with the current contents of the G_STYLES dictionary
def save_config_file():
    with open(config_file_sys, "w+", encoding="utf8") as json_file:
        json.dump(G_STYLES, json_file, indent=2)


# Called when the Close button is pressed in the dialog
def cmd_close_main_dialog(event=None):
    btn_close = CreateScriptService("DialogEvent", event)
    dialog = btn_close.Parent
    # Effectively close the dialog
    # Using XDialogView for compatibility with LO 7.3
    # dialog.endExecute(0)
    dialog.XDialogView.endExecute()


# Returns a dictionary where the translated style names are the keys and the
# style IDs are the content
def get_translated_style_names():
    translated_names = dict()
    for key in G_STYLES.keys():
        localized_name = G_STYLES[key]["localized-name"]
        translated_names[localized_name] = key
    return translated_names


# Translate a single default style name (not used for custom styles)
def translate_style_name(style_name):
    # Words to be translated in default style names
    strings_to_replace = [("Black borders", "Style.BlackBorders"), ("Variation", "Style.Variation"),
                          ("Light Blue", "Style.LightBlue"), ("Light Gray", "Style.LightGray"),
                          ("Blank", "Style.Blank"), ("Blue", "Style.Blue"),
                          ("Gray", "Style.Gray"), ("Green", "Style.Green"),
                          ("Orange", "Style.Orange"), ("Yellow", "Style.Yellow")]
    # Translates the style name
    translated_name = style_name
    for translation_pair in strings_to_replace:
        string_to_find = translation_pair[0]
        string_to_insert = l10n._(translation_pair[1])
        translated_name = translated_name.replace(string_to_find, string_to_insert)
    # Return the translated name
    return translated_name


# Translate all entries in G_STYLES and save the settings file if any translations were made
def translate_all_style_names():
    # Current locale
    try:
        # Will work starting with LO 7.4
        cur_locale = platform.OfficeLocale
    except Exception:
        cur_locale = platform.Locale
    # Indicates if changes were maded to style names
    changes_made = False
    # Check all keys to see if they need translation
    for key in G_STYLES.keys():
        if G_STYLES[key]["custom"] == 0 and G_STYLES[key]["locale"] != cur_locale:
            translated_name = translate_style_name(key)
            G_STYLES[key]["localized-name"] = translated_name
            G_STYLES[key]["locale"] = cur_locale
            changes_made = True
    # Save settings if changes were made
    if changes_made:
        save_config_file()
    # Indicates that no translations are required in the current session
    global need_style_translation
    need_style_translation = False


# Update the styles in list box using the global G_STYLES
def update_styles_list_box(list_box, selected=""):
    # List box model
    list_model = list_box.XControlModel
    # Clears all entries in the list box
    list_model.removeAllItems()
    # Insert all styles in the list box
    global G_STYLES, previews_path_url
    styles_name_dict = get_translated_style_names()
    style_names = sorted(styles_name_dict.keys())
    style_names = natural_sort(style_names)
    # First insert the favorite entries
    insert_index = 0
    # Stores the selected index based on the "selected" argument
    selected_index = 0
    for style_name in style_names:
        style_id = styles_name_dict[style_name]
        is_favorite = G_STYLES[style_id]["favorite"]
        if is_favorite == 1:
            # img_file_url = fs.BuildPath(previews_path_url, G_STYLES[style]["png-file"])
            # list_model.insertItem(insert_index, "* " + style, img_file_url)
            # style_name = G_STYLES[style]["localized-name"]
            list_model.insertItemText(insert_index, "* " + style_name)
            list_model.setItemData(insert_index, style_id)
            if style_id == selected:
                selected_index = insert_index
            insert_index += 1
    # Now insert the non-favorite entries
    for style_name in style_names:
        style_id = styles_name_dict[style_name]
        is_favorite = G_STYLES[style_id]["favorite"]
        if is_favorite == 0:
            # img_file_url = fs.BuildPath(previews_path_url, G_STYLES[style]["png-file"])
            # list_model.insertItem(insert_index, style, img_file_url)
            # style_name = G_STYLES[style]["localized-name"]
            list_model.insertItemText(insert_index, style_name)
            list_model.setItemData(insert_index, style_id)
            if style_id == selected:
                selected_index = insert_index
            insert_index += 1
    # Select the item defined in "selected"
    list_box.ListIndex = selected_index


# Returns a dictionary with the default checkbox options
def get_dlg_default_options():
    dlg_options = dict()
    dlg_options["opt_header"] = 1
    dlg_options["opt_banded"] = 1
    dlg_options["opt_font"] = 1
    dlg_options["opt_text_align"] = 1
    dlg_options["opt_compact"] = 0
    dlg_options["opt_shadow"] = 0
    dlg_options["opt_close"] = 0
    return dlg_options


# Returns a dictionary with the current state of the checkboxes in the dialog
def get_dlg_options_state(dialog):
    dlg_options = dict()
    dlg_options["opt_header"] = dialog.Controls("chkHeader").Value
    dlg_options["opt_banded"] = dialog.Controls("chkBanded").Value
    dlg_options["opt_font"] = dialog.Controls("chkFont").Value
    dlg_options["opt_text_align"] = dialog.Controls("chkTextAlign").Value
    dlg_options["opt_compact"] = dialog.Controls("chkCompact").Value
    dlg_options["opt_shadow"] = dialog.Controls("chkShadow").Value
    dlg_options["opt_close"] = dialog.Controls("chkClose").Value
    return dlg_options


# Updates the options selected in the dialog based on G_DLG_OPTIONS
def update_dlg_options_state(dialog):
    global G_DLG_OPTIONS
    if G_DLG_OPTIONS is None:
        return
    dialog.Controls("chkHeader").Value = G_DLG_OPTIONS["opt_header"]
    dialog.Controls("chkBanded").Value = G_DLG_OPTIONS["opt_banded"]
    dialog.Controls("chkFont").Value = G_DLG_OPTIONS["opt_font"]
    dialog.Controls("chkTextAlign").Value = G_DLG_OPTIONS["opt_text_align"]
    dialog.Controls("chkCompact").Value = G_DLG_OPTIONS["opt_compact"]
    dialog.Controls("chkShadow").Value = G_DLG_OPTIONS["opt_shadow"]
    dialog.Controls("chkClose").Value = G_DLG_OPTIONS["opt_close"]


# Update the dialog options from the Description string
# If the string is invalid, loads the last dialog state
def update_dlg_options_from_string(dialog, description_string):
    # Split the string values
    global G_DLG_OPTIONS
    str_values = description_string.strip().split("|")
    if len(str_values) != 7:
        update_dlg_options_state(dialog)
        return
    for value in str_values[1:]:
        if value not in ("0", "1"):
            update_dlg_options_state(dialog)
            return
    # If it gets here, the string is valid
    dialog.Controls("chkHeader").Value = int(str_values[1])
    dialog.Controls("chkBanded").Value = int(str_values[2])
    dialog.Controls("chkFont").Value = int(str_values[3])
    dialog.Controls("chkTextAlign").Value = int(str_values[4])
    dialog.Controls("chkCompact").Value = int(str_values[5])
    dialog.Controls("chkShadow").Value = int(str_values[6])
    # Updates the global options dictionary
    G_DLG_OPTIONS = get_dlg_options_state(dialog)


# Load localized strings onto the Main dialog
def localize_main_dialog(dialog):
    dialog.Caption = l10n._("MainDialog.Title")
    dialog.Controls("labSelectStyle").Caption = l10n._("MainDialog.SelectStyle")
    dialog.Controls("labStylePreview").Caption = l10n._("MainDialog.StylePreview")
    dialog.Controls("labStyleOptions").Caption = l10n._("MainDialog.StyleOptions")
    dialog.Controls("labAdditionalOptions").Caption = l10n._("MainDialog.AdditionalOptions")
    dialog.Controls("btnFavorite").Caption = l10n._("MainDialog.Favorite")
    dialog.Controls("btnDelete").Caption = l10n._("MainDialog.Delete")
    dialog.Controls("btnApply").Caption = l10n._("MainDialog.Apply")
    dialog.Controls("btnNewStyle").Caption = l10n._("MainDialog.NewStyle")
    dialog.Controls("btnClose").Caption = l10n._("MainDialog.Close")
    dialog.Controls("btnHelp").Caption = l10n._("MainDialog.Help")
    dialog.Controls("chkHeader").Caption = l10n._("MainDialog.HeaderRow")
    dialog.Controls("chkBanded").Caption = l10n._("MainDialog.Banded")
    dialog.Controls("chkFont").Caption = l10n._("MainDialog.Font")
    dialog.Controls("chkTextAlign").Caption = l10n._("MainDialog.Alignment")
    dialog.Controls("chkCompact").Caption = l10n._("MainDialog.Height")
    dialog.Controls("chkShadow").Caption = l10n._("MainDialog.Shadow")
    dialog.Controls("chkClose").Caption = l10n._("MainDialog.CloseAfter")


# Load localized strings onto the Main dialog
def localize_new_style_dialog(dialog):
    dialog.Caption = l10n._("NewStyle.Title")
    dialog.Controls("labName").Caption = l10n._("NewStyle.Name")
    dialog.Controls("labPreview").Caption = l10n._("NewStyle.Preview")
    dialog.Controls("btnCreate").Caption = l10n._("NewStyle.Create")
    dialog.Controls("btnCancel").Caption = l10n._("NewStyle.Cancel")


# Add an ActionListener to a dialog Button OnMouseButtonPressed event
def add_action_listener(dialog, control_name, command):
    control = dialog.Controls(control_name)
    listener = ActionListener(command)
    control.XControlView.addActionListener(listener)
    return listener


# Remove an ActionListener from a dialog Button
def remove_action_listener(dialog, control_name, listener):
    control = dialog.Controls(control_name)
    control.XControlView.removeActionListener(listener)


# Add an ItemListener to a ListBox
def add_item_listener(dialog, control_name, command):
    control = dialog.Controls(control_name)
    listener = ItemChangeListener(command)
    control.XControlView.addItemListener(listener)
    return listener


# Remove an ItemListener from a list box
def remove_item_listener(dialog, control_name, listener):
    control = dialog.Controls(control_name)
    control.XControlView.removeItemListener(listener)


# Add a KeyListener to a text box (for now there's just one)
def add_key_listener(dialog, control_name):
    control = dialog.Controls(control_name)
    listener = KeyPressedListener()
    control.XControlView.addKeyListener(listener)
    return listener


# Remove a KeyListener from a text box
def remove_key_listener(dialog, control_name, listener):
    control = dialog.Controls(control_name)
    control.XControlView.removeKeyListener(listener)


# Open the main dialog
def cmd_open_main_dialog(event=None):
    # Validate current selection (only a single table can be selected)
    # doc = XSCRIPTCONTEXT.getDocument()
    # desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    # doc = desktop.getCurrentComponent()
    # ui = CreateScriptService("UI")
    doc = bas.ThisComponent
    is_table, msg = validate_selection(doc)
    # If the current selection is invalid, an error message is displayed and nothing happens
    if not is_table:
        bas.MsgBox(l10n._("Msg.Error") + ": " + msg, bas.MB_ICONEXCLAMATION + bas.MB_OK)
        return
    # By default FileNaming is set to URL when the dialog is opened
    fs.FileNaming = "URL"
    # Gets the main dialog object
    main_dialog = CreateScriptService("Dialog", "GlobalScope", "TableFormatter", "MainDialog")
    # Localize the Main dialog
    localize_main_dialog(main_dialog)
    # Gets the styles list box
    styles_list = main_dialog.Controls("StyleList")
    # Reads all styles into the global dictionary
    global G_STYLES, G_DLG_OPTIONS
    global styles_file_sys, previews_path_url
    with open(styles_file_sys, "r", encoding="utf8") as f:
        G_STYLES = json.load(f)
    # Translates all style names if necessary
    global need_style_translation
    if need_style_translation:
        translate_all_style_names()
    # Get the style applied to the selected table (uses Description field of the table Shape)
    current_style = ""
    table_shape = doc.getCurrentSelection().getByIndex(0)
    current_style = get_name_from_string(table_shape.Description)
    # If this is the first time the dialog opens, saves the default settings
    if G_DLG_OPTIONS is None:
        # This happens when a table has no valid style in the first time the dialog opens
        G_DLG_OPTIONS = get_dlg_options_state(main_dialog)
    # Update the selected checkboxes based on the current table description
    update_dlg_options_from_string(main_dialog, table_shape.Description)
    # Fills the list box with all available styles
    update_styles_list_box(styles_list, current_style)
    # The first item is always selected when the dialog opens
    selected_index = styles_list.ListIndex
    selected_style = styles_list.XControlModel.getItemData(selected_index)
    # Initialize the image preview
    image_file_url = fs.BuildPath(previews_path_url, G_STYLES[selected_style]["png-file"])
    img_control = main_dialog.Controls("imgPreview")
    img_control.Picture = image_file_url
    img_control_model = img_control.XControlModel
    img_control_model.ScaleMode = ImageScaleMode.ISOTROPIC
    # Update the caption of the Favorite button
    is_favorite = G_STYLES[selected_style]["favorite"]
    btn_favorite = main_dialog.Controls("btnFavorite")
    if is_favorite == 1:
        btn_favorite.Caption = l10n._("MainDialog.NotFavorite")
    else:
        btn_favorite.Caption = l10n._("MainDialog.Favorite")
    # Add all Listeners to the dialog
    ls_bnt_apply = add_action_listener(main_dialog, "btnApply", "apply_style")
    ls_bnt_new_style = add_action_listener(main_dialog, "btnNewStyle", "open_new_style_dialog")
    ls_btn_close = add_action_listener(main_dialog, "btnClose", "close_main_dialog")
    ls_btn_help = add_action_listener(main_dialog, "btnHelp", "show_help")
    ls_btn_favorite = add_action_listener(main_dialog, "btnFavorite", "toggle_favorite")
    ls_btn_delete = add_action_listener(main_dialog, "btnDelete", "delete_style")
    ls_list_box = add_item_listener(main_dialog, "StyleList", "update_image")
    ls_chk_header = add_item_listener(main_dialog, "chkHeader", "update_dlg_status")
    ls_chk_banded = add_item_listener(main_dialog, "chkBanded", "update_dlg_status")
    ls_chk_font = add_item_listener(main_dialog, "chkFont", "update_dlg_status")
    ls_chk_text_align = add_item_listener(main_dialog, "chkTextAlign", "update_dlg_status")
    ls_chk_compact = add_item_listener(main_dialog, "chkCompact", "update_dlg_status")
    ls_chk_shadow = add_item_listener(main_dialog, "chkShadow", "update_dlg_status")
    ls_chk_close = add_item_listener(main_dialog, "chkClose", "update_dlg_status")
    # Open the dialog in modal mode (safer and more predictable than non-modal)
    main_dialog.Execute()
    # Save the current state of the checkboxes
    G_DLG_OPTIONS = get_dlg_options_state(main_dialog)
    # Save the configuration file if necessary
    global G_SAVE_CONFIG
    if G_SAVE_CONFIG:
        save_config_file()
        G_SAVE_CONFIG = False
    # Remove all listeners from the dialog
    remove_action_listener(main_dialog, "btnApply", ls_bnt_apply)
    remove_action_listener(main_dialog, "btnNewStyle", ls_bnt_new_style)
    remove_action_listener(main_dialog, "btnClose", ls_btn_close)
    remove_action_listener(main_dialog, "btnHelp", ls_btn_help)
    remove_action_listener(main_dialog, "btnFavorite", ls_btn_favorite)
    remove_action_listener(main_dialog, "btnDelete", ls_btn_delete)
    remove_item_listener(main_dialog, "StyleList", ls_list_box)
    remove_item_listener(main_dialog, "chkHeader", ls_chk_header)
    remove_item_listener(main_dialog, "chkBanded", ls_chk_banded)
    remove_item_listener(main_dialog, "chkFont", ls_chk_font)
    remove_item_listener(main_dialog, "chkTextAlign", ls_chk_text_align)
    remove_item_listener(main_dialog, "chkCompact", ls_chk_compact)
    remove_item_listener(main_dialog, "chkShadow", ls_chk_shadow)
    remove_item_listener(main_dialog, "chkClose", ls_chk_close)
    # Terminates the main dialog
    main_dialog.Terminate()


# Toggle the favorite status of a style (does not save config file)
def cmd_toggle_favorite(event=None):
    btn_control = CreateScriptService("DialogEvent", event)
    main_dialog = btn_control.Parent
    list_box = main_dialog.Controls("StyleList")
    selected_index = list_box.ListIndex
    selected_style = list_box.XControlModel.getItemData(selected_index)
    is_favorite = G_STYLES[selected_style]["favorite"]
    # Invert the favorite status
    if is_favorite == 1:
        G_STYLES[selected_style]["favorite"] = 0
    else:
        G_STYLES[selected_style]["favorite"] = 1
    # Update the list box
    update_styles_list_box(list_box, selected_style)
    # Update the caption of the Favorite button
    is_favorite = G_STYLES[selected_style]["favorite"]
    btn_favorite = main_dialog.Controls("btnFavorite")
    if is_favorite == 1:
        btn_favorite.Caption = l10n._("MainDialog.NotFavorite")
    else:
        btn_favorite.Caption = l10n._("MainDialog.Favorite")
    # Indicate that the config file needs saving
    global G_SAVE_CONFIG
    G_SAVE_CONFIG = True


# This is run when the selection changes in the listbox
def update_image_preview(event=None):
    styles_list = CreateScriptService("DialogEvent", event)
    selected_index = styles_list.ListIndex
    list_model = styles_list.XControlModel
    # The actual style name is stored in the data of the item (see XItemList interface)
    selected_style = list_model.getItemData(selected_index)
    global previews_path_url
    image_file = fs.BuildPath(previews_path_url, G_STYLES[selected_style]["png-file"])
    img_control = styles_list.Parent.Controls("imgPreview")
    img_control.Picture = image_file
    # Update the caption of the Favorite button
    is_favorite = G_STYLES[selected_style]["favorite"]
    btn_favorite = styles_list.Parent.Controls("btnFavorite")
    if is_favorite == 1:
        btn_favorite.Caption = l10n._("MainDialog.NotFavorite")
    else:
        btn_favorite.Caption = l10n._("MainDialog.Favorite")


# Deletes the selected style
def cmd_delete_style(event=None):
    btn_delete = CreateScriptService("DialogEvent", event)
    main_dialog = btn_delete.Parent
    # Gets the selected style ID in the list box
    list_box = main_dialog.Controls("StyleList")
    selected_index = list_box.ListIndex
    list_model = list_box.XControlModel
    style_id = list_model.getItemData(selected_index)
    # If there is just one style it is not possible to delete
    if list_model.ItemCount == 1:
        bas.MsgBox(l10n._("Msg.OneStyle_A") + "\n" + l10n._("Msg.OneStyle_B"),
                   bas.MB_ICONSTOP + bas.MB_OK)
        # Nothing else happens and exit
        return
    # Deletes the json and png files
    json_file = fs.BuildPath(styles_path_url, G_STYLES[style_id]["json-file"])
    png_file = fs.BuildPath(previews_path_url, G_STYLES[style_id]["png-file"])
    fs.DeleteFile(json_file)
    fs.DeleteFile(png_file)
    # Delete the entry from the styles dictionary
    G_STYLES.pop(style_id)
    # Save the config file and update the list box
    save_config_file()
    # Gets the next style ID in the list box (if it exists, or else gets previous)
    if selected_index < list_model.ItemCount - 1:
        new_selection = list_model.getItemData(selected_index + 1)
    else:
        new_selection = list_model.getItemData(selected_index - 1)
    update_styles_list_box(list_box, new_selection)
    # Update the preview image
    image_file_url = fs.BuildPath(previews_path_url, G_STYLES[new_selection]["png-file"])
    img_control = main_dialog.Controls("imgPreview")
    img_control.Picture = image_file_url


"""
---------------------------------------------------------------------------------------------------
Code used to implement the NewStyleDialog
---------------------------------------------------------------------------------------------------
"""


# Opens the dialog if the table is at least 3 x 3 and captures the current preview
def cmd_open_new_style_dialog(event=None):
    # Gets the selected table Shape and Model
    # doc = XSCRIPTCONTEXT.getDocument()
    doc = bas.ThisComponent
    table_shape = doc.getCurrentSelection().getByIndex(0)
    table_model = table_shape.Model
    # Check if the table size is sufficient to extract the style
    nRows = table_model.Rows.getCount()
    nCols = table_model.Columns.getCount()
    if nCols < 3 or nRows < 5:
        bas.MsgBox(l10n._("Msg.Error") + ": " + l10n._("Msg.NewStyle"),
                   bas.MB_ICONEXCLAMATION + bas.MB_OK)
        # If the selection is invalid, return without opening
        return
    # Opens the New Style dialog and shows the preview file
    try:
        # Creates the preview file
        global current_json_temp_url, current_png_temp_url
        current_json_temp_url, current_png_temp_url = create_temp_json_and_png()
        # Creates the dialog and opens it
        dialog = CreateScriptService("Dialog", "GlobalScope", "TableFormatter", "NewStyleDialog")
        img_control = dialog.Controls("imgPreview")
        img_control.Picture = current_png_temp_url
        img_control_model = img_control.XControlModel
        img_control_model.ScaleMode = ImageScaleMode.ISOTROPIC
        # Localize the New Style dialog
        localize_new_style_dialog(dialog)
        # Assign listeners to all controls in the dialog
        lst_create = add_action_listener(dialog, "btnCreate", "create-new-style")
        lst_new_style = add_action_listener(dialog, "btnCancel", "close-new-style-dialog")
        lst_textbox = add_key_listener(dialog, "textName")
        # Gets main_dialog to be able to center the child dialog
        btn_parent = CreateScriptService("DialogEvent", event)
        main_dialog = btn_parent.Parent
        # Centers child dialog
        try:
            # Only available from LO 7.4 onwards
            dialog.Center(main_dialog)
        except Exception:
            pass
        dialog.Execute()
        # Removes all listeners
        remove_action_listener(dialog, "btnCreate", lst_create)
        remove_action_listener(dialog, "btnCancel", lst_new_style)
        remove_key_listener(dialog, "textName", lst_textbox)
        # Terminates the dialog object
        dialog.Terminate()
        # Checks if the main dialog needs to reload the list box
        global b_needs_reloading, s_new_style_name
        if b_needs_reloading:
            styles_list_box = main_dialog.Controls("StyleList")
            update_styles_list_box(styles_list_box, s_new_style_name)
            image_file_url = fs.BuildPath(previews_path_url, G_STYLES[s_new_style_name]["png-file"])
            img_control = main_dialog.Controls("imgPreview")
            img_control.Picture = image_file_url
            b_needs_reloading = False
            s_new_style_name = ""
            # The newly created style is not a favorite; update the Caption of btnFavorite
            btn_favorite = main_dialog.Controls("btnFavorite")
            btn_favorite.Caption = l10n._("MainDialog.Favorite")
    except Exception as e:
        # If all is correct, this will never occur
        bas.MsgBox(l10n._("Msg.PreviewError") + "\n\n" + str(e),
                   bas.MB_ICONSTOP + bas.MB_OK)


# Saves the newly created style
def cmd_save_new_style(event=None):
    # Gets the name given by the user to the new style
    btn_save = CreateScriptService("DialogEvent", event)
    dialog = btn_save.Parent
    text_name = dialog.Controls("textName")
    new_style_name = text_name.Value.strip()
    # Check if the style name already exists
    global G_STYLES
    if new_style_name in G_STYLES:
        bas.MsgBox(l10n._("Msg.StyleExists_A", new_style_name) + "\n" +
                   l10n._("Msg.StyleExists_B"),
                   bas.MB_ICONSTOP + bas.MB_OK)
        # Does not save the style
        return
    elif new_style_name == "":
        # If an empty string is provided, an error occurs
        bas.MsgBox(l10n._("Msg.EmptyName"),
                   bas.MB_ICONSTOP + bas.MB_OK)
        # Does not save the style
        return
    # Copy the newly created style
    global current_json_temp_url, current_png_temp_url
    new_file_basename = get_valid_style_filename()
    new_json_file = fs.BuildPath(styles_path_url, new_file_basename + ".json")
    new_png_file = fs.BuildPath(previews_path_url, new_file_basename + ".png")
    fs.MoveFile(current_json_temp_url, new_json_file)
    fs.MoveFile(current_png_temp_url, new_png_file)
    # Update the G_STYLES dictionary
    # Current locale
    try:
        # Will work starting with LO 7.4
        cur_locale = platform.OfficeLocale
    except Exception:
        cur_locale = platform.Locale
    new_style_props = dict()
    new_style_props["localized-name"] = new_style_name
    new_style_props["locale"] = cur_locale
    new_style_props["json-file"] = new_file_basename + ".json"
    new_style_props["png-file"] = new_file_basename + ".png"
    new_style_props["favorite"] = 0
    new_style_props["custom"] = 1
    G_STYLES[new_style_name] = new_style_props
    # Update the config file and close the dialog
    save_config_file()
    # Using XDialogView for compatibility with LO 7.3
    # dialog.endExecute(0)
    dialog.XDialogView.endExecute()
    # Indicate that the styles list box needs reloading
    global b_needs_reloading, s_new_style_name
    b_needs_reloading = True
    s_new_style_name = new_style_name


# Detects when the user presses ENTER when the cursor is in the edit box
def on_new_style_key_pressed(event=None):
    if event.KeyCode == Key.RETURN:
        # Calls the procedure to save the new style
        cmd_save_new_style(event)


# Closes the New style dialog
def cmd_close_new_style_dialog(event=None):
    # Delete the temporary files created
    global current_json_temp_url, current_png_temp_url
    fs.DeleteFile(current_json_temp_url)
    fs.DeleteFile(current_png_temp_url)
    # Close the dialog
    btn_close = CreateScriptService("DialogEvent", event)
    dialog = btn_close.Parent
    # Using XDialogView for compatibility with LO 7.3
    # dialog.endExecute(0)
    dialog.XDialogView.endExecute()


# Returns a file name that does not exist using the custom_style_ prefix
# Returns only the base name (not the path nor the extension)
def get_valid_style_filename():
    global styles_path_url
    file_prefix = "custom_style_{}"
    # Tests up to CUSTOM_STYLE_LIMIT (1000) custom styles (hope no one ever gets there)
    for i in range(CUSTOM_STYLE_LIMIT):
        new_file_name = file_prefix.format(i)
        path_to_file = fs.BuildPath(styles_path_url, new_file_name + ".json")
        if not fs.FileExists(path_to_file):
            return new_file_name
    # If it reaches here it is because there are more than CUSTOM_STYLE_LIMIT custom styles
    bas.MsgBox(l10n._("Msg.StyleLimit") + "\n",
               bas.MB_ICONSTOP + bas.MB_OK)


# Create a temporary file name
def create_temp_file_name():
    file_num = rnd.randint(0, 100000)
    base_name = f"temp_{file_num:05d}"
    file_name = fs.BuildPath(temp_folder_url, base_name)
    return file_name


# Creates a temporary Json file and PNG image of the selected table
def create_temp_json_and_png(args=None):
    # Gets the selected table
    this_doc = bas.ThisComponent
    table_shape = this_doc.getCurrentSelection().getByIndex(0)
    selected_table = table_shape.Model
    # Creates a temporary file name for the Json and PNG files
    json_file_url = create_temp_file_name() + ".json"
    json_file_sys = bas.ConvertFromUrl(json_file_url)
    png_file_url = create_temp_file_name() + ".png"
    # Create the Json file from selected table
    create_json_from_table(selected_table, json_file_sys)
    # Open the template file in hidden mode
    ui = CreateScriptService("UI")
    global template_odp_url
    template_doc = ui.OpenDocument(template_odp_url, hidden=True)
    doc_component = template_doc.XComponent
    first_slide = doc_component.DrawPages.getByIndex(0)
    table_shape = first_slide.getByIndex(0)
    table_model = table_shape.Model
    apply_format_to_table(table_model, json_file_sys, get_dlg_default_options())
    # Create the PNG preview
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    export_filter = smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", ctx)
    prop_dict = dict()
    prop_dict["URL"] = png_file_url
    prop_dict["MimeType"] = "image/png"
    sf_dict = CreateScriptService("Dictionary", prop_dict)
    filter_args = sf_dict.ConvertToPropertyValues()
    export_filter.setSourceDocument(table_shape)
    export_filter.filter(filter_args)
    # Close the template document
    template_doc.CloseDocument(saveask=False)
    # Return the files created (both in URL format)
    return json_file_url, png_file_url


"""
---------------------------------------------------------------------------------------------------
Action listener used in all buttons
---------------------------------------------------------------------------------------------------
"""


class ActionListener(unohelper.Base, XActionListener):
    def __init__(self, command):
        self.command = command

    def actionPerformed(self, event):
        if self.command == "apply_style":
            cmd_apply_pressed(event)
        elif self.command == "close_main_dialog":
            cmd_close_main_dialog(event)
        elif self.command == "open_new_style_dialog":
            cmd_open_new_style_dialog(event)
        elif self.command == "show_help":
            pass
        elif self.command == "toggle_favorite":
            cmd_toggle_favorite(event)
        elif self.command == "delete_style":
            cmd_delete_style(event)
        elif self.command == "create-new-style":
            cmd_save_new_style(event)
        elif self.command == "close-new-style-dialog":
            cmd_close_new_style_dialog(event)
        else:
            print("Some error happened while triggering an ActionListener")
            pass

    def dispose(self, event):
        pass


"""
---------------------------------------------------------------------------------------------------
Action listener used in the style list box and in all checkboxes
---------------------------------------------------------------------------------------------------
"""


class ItemChangeListener(unohelper.Base, XItemListener):
    def __init__(self, command):
        self.command = command

    def itemStateChanged(self, event):
        if self.command == "update_image":
            update_image_preview(event)
        elif self.command == "update_dlg_status":
            on_checkbox_click(event)
        else:
            pass

    def dispose(self, event):
        pass


"""
---------------------------------------------------------------------------------------------------
Export component to be launched by LibreOffice using XJobExecutor
---------------------------------------------------------------------------------------------------
"""


class KeyPressedListener(unohelper.Base, XKeyListener):
    def __init__(self):
        pass

    def keyPressed(self, event):
        on_new_style_key_pressed(event)

    def keyReleased(self, event):
        pass


"""
---------------------------------------------------------------------------------------------------
Export component to be launched by LibreOffice using XJobExecutor
---------------------------------------------------------------------------------------------------
"""


class TableFormatter(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        # Initialize ScriptForge services
        global bas, exc, fs, platform
        bas = CreateScriptService("Basic")
        exc = CreateScriptService("Exception")
        fs = CreateScriptService("FileSystem")
        platform = CreateScriptService("Platform")
        # Initialize the extension paths and L10n
        global b_needs_initialize_l10n
        if b_needs_initialize_l10n:
            initialize_paths()
            initialize_l10n()
            b_needs_initialize_l10n = False
        # Open the main dialog
        cmd_open_main_dialog(self.ctx)


# Export implementation
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    TableFormatter, "rafael.lima.TableFormatter",
    ("com.sun.star.Job", ), )
