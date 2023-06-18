#' AMReader
#'
#' AMReader takes data collected in AMScorer, processes
#' it, performs statistical analyses and enables production and customisation of
#' a graphical output. Please see README for further information.
#'
#' See the README for further information.
#'
#' @param Experiment_Name A string with the experiment name used to set the
#' defaults for File_Name, Stat_File, and Graph_File.
#'
#' @param Path A string giving the path to the input excel file.
#'
#' @param File_Name A string giving the name of the input excel file.
#'
#' @param Condition_Include A vector of characters which may take values from
#' A to Z, AA to AZ and BA to BZ.
#'
#' @param Condition_Exclude A vector of characters which may take values from
#' A to Z, AA to AZ and BA to BZ.
#'
#' @param Structure_Include A vector of characters which may take values from
#' "Total", "EH", "H", "IH", "A", "V" or "S".
#'
#' @param Structure_Exclude A vector of characters which may take values from
#' "Total", "EH", "H", "IH", "A", "V" or "S".
#'
#' @param Facet_1 A logical: If TRUE, data from Facet_1 will be processed.
#'
#' @param Facet_2 A logical: If TRUE, data from Facet_2 will be processed.
#'
#' @param Facet_3 A logical: If TRUE, data from Facet_3 will be processed.
#'
#' @param Stat_Dunn_Padj A string which may take a value from
#' "none", "bonferroni", "holm", "hommel", "hochberg", "BH", "BY" or "fdr".
#'
#' @param Stat_Wilcoxon_Padj A string which may take a value from
#' "none", "bonferroni", "holm", "hommel", "hochberg", "BH", "BY" or "fdr".
#'
#' @param Stat_Sided A string which may take a value from "two.sided", "less",
#' or "greater".
#'
#' @param Stat_Output A logical: If TRUE, the Statistical Analysis output will
#' be saved to excel
#'
#' @param Stat_File A string giving the file name of the output from the
#' statistical analysis file. Does not include ".xlsx".
#'
#' @param Graph_Type A string which may take a value from "Facets" or "Single".
#'
#' @param Graph_Sample_Sizes A logical: If TRUE, sample size information will be
#' displayed on the graph.
#'
#' @param Graph_Condition_Order A vector of characters which may take values
#' from A to Z, AA to AZ and BA to BZ.
#'
#' @param Graph_Facet_1_Order A vector of characters giving the possible values
#' of Facet_1 in the "Conditions" tab of the associated excel input file.
#'
#' @param Graph_Facet_2_Order A vector of characters giving the possible values
#' of Facet_2 in the "Conditions" tab of the associated excel input file.
#'
#' @param Graph_Facet_3_Order A vector of characters giving the possible values
#' of Facet_3 in the "Conditions" tab of the associated excel input file.
#'
#' @param Graph_Stat_Test A string which may take a value from
#' "Wilcoxon", "Tukey", "Dunn", or "NULL".
#'
#' @param Graph_Stat_Display A string which may take a value from
#' "Letters" or "Reference"
#'
#' @param Graph_Reference_Condition A string which may take a values from
#' "Letters" or "Reference"
#'
#' @param Graph_Manual_Colour A logical: If TRUE, the input from Manual_Colour
#' given in the "Conditions" tab of the associated excel input file.
#'
#' @param Graph_Colour A string which may take a values from
#' "Viridis", "Brewer", "Grey"
#'
#' @param Graph_Palette A string which may take any value which defines a valid
#' palette from Viridis or Brewer colour schemes. See the R documentation
#' associated with both colour schemes.
#'
#' @param Graph_Palette_Begin A numeric taking values between 0 and 1.
#'
#' @param Graph_Palette_End A numeric taking values between 0 and 1.
#'
#' @param Graph_Text_Colour A string which may take any valid colour input
#' listed in color() or any hexcode.
#'
#' @param Graph_Background_Colour A string which may take any valid colour input
#' listed in color() or any hexcode.
#'
#' @param Graph_Hline_Colour A string which may take any valid colour input
#' listed in color() or any hexcode.
#'
#' @param Graph_Legend A logical: If FALSE, the graph legend will not be
#' displayed
#'
#' @param Graph_Size_Right_Label A numeric value greater than 0. Multiplies
#' the size of the "Structure" labels.
#'
#' @param Graph_Size_Top_Label A numeric value greater than 0. Multiplies
#' the size of the "Facet" labels.
#'
#' @param Graph_Size_Y_Axis A numeric value greater than 0. Multiplies
#' the size of the "Y_axis" numbers
#'
#' @param Graph_Size_X_Axis A numeric value greater than 0. Multiplies
#' the size of the "X_axis" text.
#'
#' @param Graph_Size_Legend A numeric value greater than 0. Multiplies
#' the size of the figure legend.
#'
#' @param Graph_Size_Legend_Text A numeric value greater than 0. Multiplies
#' the size of the "Legend" text.
#'
#' @param Graph_Size_Percentages A numeric value greater than 0. Multiplies
#' the size of the Y-label "Percentage Colonisation" text.
#'
#' @param Graph_Size_Statistics A numeric value greater than 0. Multiplies
#' the size of the "statistical information" text.
#'
#' @param Graph_Size_Datapoints A numeric value greater than 0. Multiplies
#' the size of the "data points".
#'
#' @param Graph_Output A logical: If TRUE, the graph will be saved as an
#' output
#'
#' @param Graph_Resolution A string giving the desired dpi of the output
#' graph file
#'
#' @param Graph_Width_Adjustment A string giving the desired numeric
#' for stretching the graph output's width. Use only if necessary.
#'
#' @param Graph_Height_Adjustment A string giving the desired numeric
#' for stretching the graph output's height. Use only if necessary.
#'
#' @param Graph_File A string giving the file name of the graph output.
#' Specify the file extension (.png, .pdf, .jpeg etc.)
#'
#' @import ggplot2
#' @importFrom plyr revalue
#' @importFrom glue glue
#' @import dplyr
#' @importFrom readxl read_excel
#' @importFrom tidyr pivot_longer
#' @import openxlsx
#' @importFrom magrittr %>%
#' @importFrom purrr map_dfr
#' @importFrom multcompView multcompLetters
#' @importFrom tibble rownames_to_column
#' @importFrom rcompanion cldList
#' @importFrom rcompanion fullPTable
#' @importFrom broom tidy
#' @importFrom DescTools DunnTest
#'
#' @export

AMReader <- function(
    Experiment_Name = "",
    Path = getwd(),
    File_Name = "AMScorer {Experiment_Name}.xlsx",
    Condition_Include = NULL,
    Condition_Exclude = NULL,
    Structure_Include = NULL,
    Structure_Exclude = NULL,
    Facet_1 = FALSE,
    Facet_2 = FALSE,
    Facet_3 = FALSE,
    Stat_Dunn_Padj = "none",
    Stat_Wilcoxon_Padj = "none",
    Stat_Sided = "two.sided",
    Stat_Output = FALSE,
    Stat_File = "Statistical Analysis {Experiment_Name}.xlsx",
    Graph_Type = "Facets",
    Graph_Sample_Sizes = NULL,
    Graph_Condition_Order = NULL,
    Graph_Facet_1_Order = NULL,
    Graph_Facet_2_Order = NULL,
    Graph_Facet_3_Order = NULL,
    Graph_Stat_Test = NULL,
    Graph_Stat_Display = "Letters",
    Graph_Reference_Condition = NULL,
    Graph_Manual_Colour = FALSE,
    Graph_Colour = "Viridis",
    Graph_Palette = "viridis",
    Graph_Palette_Begin = 0,
    Graph_Palette_End = 1,
    Graph_Text_Colour = "grey30",
    Graph_Background_Colour = "grey95",
    Graph_Hline_Colour = "grey30",
    Graph_Legend = TRUE,
    Graph_Size_Right_Label = 1,
    Graph_Size_Top_Label = 1,
    Graph_Size_Y_Axis = 1,
    Graph_Size_X_Axis = 1,
    Graph_Size_Legend = 1,
    Graph_Size_Legend_Text = 1,
    Graph_Size_Percentages = 1,
    Graph_Size_Statistics = 1,
    Graph_Size_Datapoints = 1,
    Graph_Output = FALSE,
    Graph_Resolution = 1200,
    Graph_Width_Adjustment = 1,
    Graph_Height_Adjustment = 1,
    Graph_File = "Colonisation Percentage Graph {Experiment_Name}.png")

{

  #####
  # A) Check user inputs into the function.
  #####

  # Create the data necessary for checking user inputs

  # A list of acceptable conditions is required
  # These conditions are A-Z, AA-AZ and BA-BZ
  # These are contained within Condition_Letter_List

  Condition_Letter_List <- data.frame(
    c(
      LETTERS[1:26],
      sort(apply(expand.grid(LETTERS[1:2],
                             LETTERS[1:26]),
                 MARGIN = 1,
                 paste0,
                 collapse = "")))
  ) %>%
    rename("Letters" = 1)

  # A list of acceptable Structures is required

  Structures_list <- c("Total", "EH", "H", "IH", "A", "V", "S")

  # Set the Experiment_Name variable
  # Defaults to ""

  if (missing(Experiment_Name)
      | is.null(Experiment_Name)) {
    print("Please provide an experiment name with: Experiment_Name = ")
    Experiment_Name <- ""
  } else {
    Experiment_Name <- Experiment_Name
  }

  # Set the Path variable
  # Defaults to the current working directory

  if (missing(Path)
      | is.null(Path)) {
    print("No path identified, defaulting to the current working directory")
    Path <- getwd()
  } else {
    Path <- Path
  }

  # Set the File_Name variable
  # Defaults to "AMScorer {Experiment_Name}.xlsx"
  # User can employ the file.choose() function with 'File_Name = CHOOSE'

  if (missing(File_Name)
      | is.null(File_Name)) {
    print(
      paste0(
        "No File_Name input, defaulting to ",
        "'AMScorer {Experiment_Name}.xlsx'"))
    File_Name <- paste("AMScorer ", Experiment_Name, ".xlsx",
                       sep = "")
  } else if (File_Name == "CHOOSE") {
    Chosen_Path <- file.choose()
  } else {
    File_Name <- File_Name
  }

  # If the user has not used the file.choose() function, define the...
  # Chosen_Path by combining the Path and File_Name input

  if (File_Name != "CHOOSE") {
    Chosen_Path <- paste(Path, File_Name, sep = "/")
  }

  # Save data to global environment for user

  assign("Chosen_Path", Chosen_Path, envir=.GlobalEnv)

  # Process Include/Exclude Structures/Conditions

  # Only one of Condition_Include and Condition_Exclude can be used...
  # so an error will be induced if neither are NULL or MISSING
  # If both are NULL or MISSING, the default is to include all conditions...
  # which is named Conditions_to_include
  # If Condition_Include is set by the user, this replaces Conditions_to_include
  # The input of Condition_Include and Condition_Exclude is checked to...
  # confirm that all the listed conditions are valid (A-Z, AA-AZ and BA-BZ)

  if((!missing(Condition_Include)
      & !missing(Condition_Exclude))
     & (!is.null(Condition_Include)
        & !is.null(Condition_Exclude))) {
    stop(paste0("Cannot use Condition_Include and ",
                "Condition_Exclude simultaneously - use only one"))

  } else if ((missing(Condition_Include)
              | is.null(Condition_Include))
             & (missing(Condition_Exclude)
                | is.null(Condition_Exclude))) {
    print(paste("Condition_Include and Condition_Exclude are unassigned.",
                " All conditions will appear in the output",
                sep = ""))
    Conditions_to_include <- unlist(list(Condition_Letter_List))

  } else {

    # Check that the user input from Condition_Include or Condition_Exclude...
    # is valid (contained within Condition_Letter_List)
    # If not, report to the user

    invalid_condition <- Condition_Include[!(Condition_Include %in%
                                               Condition_Letter_List$Letters)]

    if (length(invalid_condition) > 0) {
      warning(paste(
        "The following values from Condition_Include are not valid: '",
        paste0(invalid_condition, collapse = ", "),
        "' - Valid condition codes are A-Z, AA-AZ, BA-BZ. '",
        paste0(invalid_condition, collapse = ", "),
        "' will be ignored",
        sep = ""))
    }

    invalid_condition <- Condition_Exclude[!(Condition_Exclude %in%
                                               Condition_Letter_List$Letters)]

    if (length(invalid_condition) > 0) {
      warning(paste(
        "The following values from Condition_Exclude are not valid: '",
        paste0(invalid_condition, collapse = ", "),
        "' - Valid condition codes are A-Z, AA-AZ, BA-BZ. '",
        paste0(invalid_condition, collapse = ", "),
        "' will be ignored",
        sep = ""))
    }

    # Set the new variable Conditions_to_include from user input.
    # Condition_Exclude is processed later

    Conditions_to_include <- Condition_Include

  }

  # Only one of Structure_Include and Structure_Exclude can be used...
  # so an error will be induced if neither are NULL or MISSING
  # If both are NULL or MISSING, the default is to include all structures...
  # which is named Structures
  # If Structure_Include is set by the user, this becomes Structures
  # If Structure_Exclude is set by the user, a list of structures is produced...
  # by removing the user input from Structures_list, and called Structures
  # The input of Structure_Include is checked to confirm that all the listed...
  # structures are valid ("Total", "EH", "H", "IH", "A", "V", "S")

  if((!missing(Structure_Include)
      & !missing(Structure_Exclude))
     & (!is.null(Structure_Include)
        & !is.null(Structure_Exclude)))  {

    stop(paste0("Cannot use Structure_Include and Structure_Exclude ",
                "simultaneously - use only one"))

  } else if ((missing(Structure_Include)
              | is.null(Structure_Include))
             & (missing(Structure_Exclude)
                | is.null(Structure_Exclude))) {

    print(paste0("Structure_Include and Structure_Exclude are unassigned. ",
                 "All Structures will appear in the output"))

    Structures <- Structures_list

  } else {

    # Create a dataframe of acceptable structures

    Structures_list_df <- data.frame(Structures_list) %>%
      rename("Structures" = 1)

    # Check the Structures given by the user are found in Structures_list_df

    invalid_structure <- Structure_Include[!(Structure_Include %in%
                                               Structures_list_df$Structures)]

    if (length(invalid_structure) > 0) {
      stop(paste(
        "The following values from Structure_Include are not valid. '",
        paste0(invalid_structure, collapse = ", "),
        "' - Valid Structure codes are: Total, EH, H, IH, A, V, S",
        sep = ""))
    }

    invalid_structure <- Structure_Exclude[!(Structure_Exclude %in%
                                               Structures_list_df$Structures)]

    if (length(invalid_structure) > 0) {
      stop(paste(
        "The following values from Structure_Exclude are not valid. '",
        paste0(invalid_structure, collapse = ", "),
        "' - Valid Structure codes are: Total, EH, H, IH, A, V, S",
        sep = ""))
    }

    # If the user is using Structure_Include

    if (!missing(Structure_Include)
        & !is.null(Structure_Include)) {

      # Set Structures to the user input of Structure_Include

      Structures <- Structure_Include
    }

    # If the user is using Structure_Exclude

    if (!missing(Structure_Exclude)
        & !is.null(Structure_Exclude)) {

      # Set Structures to those Structures in Structures_list not excluded...
      # by the user

      Structures <- Structures_list[!Structures_list %in% Structure_Exclude]
    }

  }

  # Set the Facet_1 variable
  # Defaults to FALSE

  if (missing(Facet_1)
      | is.null(Facet_1)) {
    Facet_1 <- FALSE
  } else {
    Facet_1 <- Facet_1
  }

  # Set the Facet_2 variable
  # Defaults to FALSE

  if (missing(Facet_2)
      | is.null(Facet_2)) {
    Facet_2 <- FALSE
  } else {
    Facet_2 <- Facet_2
  }

  # Set the Facet_3 variable
  # Defaults to FALSE

  if (missing(Facet_3)
      | is.null(Facet_3)) {
    Facet_3 <- FALSE
  } else {
    Facet_3 <- Facet_3
  }

  # There are multiple options for the Stat_Dunn_Padj variable:

  Stat_Dunn_Padj_options <- c("none", "bonferroni", "hommel", "holm",
                              "hochberg", "BH", "BY", "fdr")

  # Set the Stat_Dunn_Padj variable to one of the above options
  # Defaults to "none"
  # Stat_Dunn_Padj must be one of these

  if ((missing(Stat_Dunn_Padj)
       | is.null(Stat_Dunn_Padj))) {
    print("No Stat_Dunn_Padj input, defaulting to Stat_Dunn_Padj = 'none'")
    Stat_Dunn_Padj <- "none"
  } else if (Stat_Dunn_Padj %in% Stat_Dunn_Padj_options) {
    Stat_Dunn_Padj <- Stat_Dunn_Padj
  } else {
    stop("Stat_Dunn_Padj must be one of: ", paste0(Stat_Dunn_Padj_options,
                                                   collapse = ", "))
  }

  # There are multiple options for the Stat_Wilcoxon_Padj variable:

  Stat_Wilcoxon_Padj_options <- c("none", "holm", "hochberg", "hommel",
                                  "bonferroni", "BH", "BY", "fdr")

  # Set the Stat_Wilcoxon_Padj variable to one of the above options
  # Defaults to "none"
  # Stat_Wilcoxon_Padj must be one of these

  if (missing(Stat_Wilcoxon_Padj)
      | is.null(Stat_Wilcoxon_Padj)) {
    print(paste0("No Stat_Wilcoxon_Padj input, ",
                 "defaulting to Stat_Wilcoxon_Padj = 'none'"))
    Stat_Wilcoxon_Padj <- "none"
  } else if (Stat_Wilcoxon_Padj %in% Stat_Wilcoxon_Padj_options) {
    Stat_Wilcoxon_Padj <- Stat_Wilcoxon_Padj
  } else {
    stop("Stat_Wilcoxon_Padj must be one of: ",
         paste0(Stat_Wilcoxon_Padj_options,
                collapse = ", "))
  }

  # There are multiple options for the Stat_Sided variable:

  Stat_Sided_options <- c("two.sided", "less", "greater")

  # Set the Stat_Sided variable to one of the above options
  # Defaults to "two.sided"
  # Stat_Sided must be one of these

  if ((missing(Stat_Sided)
       | is.null(Stat_Sided))) {
    print("No Stat_Sided input, defaulting to Stat_Sided = 'two.sided'")
    Stat_Sided <- "two.sided"
  } else if (Stat_Sided %in% Stat_Sided_options) {
    Stat_Sided <- Stat_Sided
  } else {
    stop("Stat_Sided must be one of: ", paste0(Stat_Sided_options,
                                               collapse = ", "))
  }

  # Set the Stat_Output variable to either TRUE or FALSE
  # Defaults to FALSE

  if (missing(Stat_Output)
      | is.null(Stat_Output)) {
    print("No Stat_Output input, defaulting to Stat_Output = FALSE")
    Stat_Output <- FALSE
  } else if ((Stat_Output == TRUE)
             | (Stat_Output == FALSE)) {
    Stat_Output <- Stat_Output
  } else {
    stop("Stat_Output must be either TRUE or FALSE")
  }

  # Set the Stat_File variable
  # Defaults to "Statistical Analysis {Experiment_Name}.xlsx"

  if (missing(Stat_File)
      | is.null(Stat_File)) {
    print(paste0("No Stat_File input, defaulting to ",
                 "Statistical Analysis {Experiment_Name}.xlsx"))
    Stat_File <- glue("Statistical Analysis {Experiment_Name}.xlsx")
    Stat_File <- gsub(x = Stat_File, " .xlsx", ".xlsx")
  } else {
    Stat_File <- paste(Stat_File, ".xlsx", sep = "")
  }

  # Set the Graph_Type variable to either "Facets" or "Single"
  # Defaults to "Facets"

  if (missing(Graph_Type)
      | is.null(Graph_Type)) {
    print("No Graph_Type input, defaulting to Facets")
    Graph_Type <- "Facets"
  } else if ((Graph_Type == "Facets")
             | (Graph_Type == "Single")) {
    Graph_Type <- Graph_Type
  } else {
    stop("Graph_Type must be either 'Facets' (default) or 'Single'")
  }

  # Set the Graph_Sample_Sizes variable to either TRUE or FALSE
  # Defaults to Graph_Sample_Sizes = FALSE

  if (missing(Graph_Sample_Sizes)
      | is.null(Graph_Sample_Sizes)) {
    print("No Graph_Sample_Sizes input, defaulting to Graph_Sample_Sizes = FALSE")
    Graph_Sample_Sizes <- FALSE
  } else if ((Graph_Sample_Sizes == TRUE)
             | (Graph_Sample_Sizes == FALSE)) {
    Graph_Sample_Sizes <- Graph_Sample_Sizes
  } else {
    stop("Graph_Sample_Sizes must be either TRUE or FALSE")
  }

  # Set the Graph_Condition_Order variable
  # Defaults to the alphabetical list A-Z, AA-AZ and BA-BZ

  if (missing(Graph_Condition_Order)
      | is.null(Graph_Condition_Order)) {
    print(paste0("No Graph_Condition_Order input, conditions will appear ",
                 "alphabetically (given by the condition letter code)"))
    Graph_Condition_Order <- Condition_Letter_List
  } else {

    # Check the conditions given by the user are found in Condition_Letter_List

    invalid_condition <-
      Graph_Condition_Order[!(Graph_Condition_Order %in%
                                Condition_Letter_List$Letters)]

    if (length(invalid_condition) > 0) {
      warning(paste(
        "The following values from Graph_Condition_Order are not valid: ",
        invalid_condition,
        " - Valid condition codes are A-Z, AA-AZ, BA-BZ. '",
        invalid_condition,
        "' will be ignored",
        sep = ""))
    }

    Graph_Condition_Order <- Graph_Condition_Order

  }

  # Set the Graph_Facet_1_Order variable
  # Defaults to NULL

  if (missing(Graph_Facet_1_Order)
      | is.null(Graph_Facet_1_Order)) {
    Graph_Facet_1_Order <- NULL
    print(paste("No Graph_Facet_1_Order input, ",
                "defaulting to Graph_Facet_1_Order = NULL",
                sep = ""))
  } else {
    Graph_Facet_1_Order <- Graph_Facet_1_Order
  }

  # Set the Graph_Facet_2_Order variable
  # Defaults to NULL

  if (missing(Graph_Facet_2_Order)
      | is.null(Graph_Facet_2_Order)) {
    Graph_Facet_2_Order <- NULL
    print(paste("No Graph_Facet_2_Order input, ",
                "defaulting to Graph_Facet_2_Order = NULL",
                sep = ""))
  } else {
    Graph_Facet_2_Order <- Graph_Facet_2_Order
  }

  # Set the Graph_Facet_3_Order variable
  # Defaults to NULL

  if (missing(Graph_Facet_3_Order)
      | is.null(Graph_Facet_3_Order)) {
    Graph_Facet_3_Order <- NULL
    print(paste("No Graph_Facet_3_Order input, ",
                "defaulting to Graph_Facet_3_Order = NULL",
                sep = ""))
  } else {
    Graph_Facet_3_Order <- Graph_Facet_3_Order
  }

  # There are 4 options for the Graph_Stat_Test_options variable:

  Graph_Stat_Test_options <- c("Wilcoxon", "Tukey", "Dunn", "NULL")

  # Set the Graph_Stat_Test variable to one of the above options
  # Defaults to NULL
  # Graph_Stat_Test must be one of these

  if (missing(Graph_Stat_Test)
      | is.null(Graph_Stat_Test)) {
    print("No Graph_Stat_Test input, defaulting to Graph_Stat_Test = NULL")
    Graph_Stat_Test <- NULL
  } else if (Graph_Stat_Test %in% Graph_Stat_Test_options) {
    Graph_Stat_Test <- Graph_Stat_Test
  } else {
    stop("Graph_Stat_Test must be one of: ", paste0(Graph_Stat_Test_options,
                                                    collapse = ", "))
  }

  # Set the Graph_Stat_Display variable to either Letters or Reference
  # Defaults to Letters
  # Graph_Stat_Display must be one of these

  if (missing(Graph_Stat_Display)
      | is.null(Graph_Stat_Display)) {
    print("No Graph_Stat_Display, defaulting to Graph_Stat_Display = 'Letters'")
    Graph_Stat_Display <- "Letters"
  } else if ((Graph_Stat_Display == "Letters")
             | (Graph_Stat_Display == "Reference")) {
    Graph_Stat_Display <- Graph_Stat_Display
  } else {
    stop("Graph_Stat_Display must be one of: 'Letters' or 'Reference'")
  }

  # Graph_Type == "Single" cannot have stats displayed over it

  if ((Graph_Type == "Single")
      & (!is.null(Graph_Stat_Test))) {
    warning(paste("Statistics cannot be displayed if Graph_Type == 'Single'.",
                  " Choose Graph_Type = 'Facets' to see statistics", sep = ""))
  }

  # Set the Graph_Reference_Condition variable
  # Defaults to NULL
  # Later, it will be checked whether this Condition appears in the data
  # If Graph_Reference_Condition is unspecified, it is not possible to use...
  # Graph_Stat_Display = "Reference". Therefore, an error will be produced
  # asking user to define the reference condition

  if ((missing(Graph_Reference_Condition)
       | is.null(Graph_Reference_Condition))) {
    print(paste("No Graph_Reference_Condition input"))
    Graph_Reference_Condition <- NULL

    if (Graph_Stat_Display == "Reference") {
      stop(paste0("Graph_Stat_Display = 'Reference', however no",
                  " Graph_Reference_Condition has been provided.",
                  " Please specify the letter code of the reference condition",
                  " using e.g. 'Graph_Reference_Condition = A'"))
    }

  } else {
    Graph_Reference_Condition <- Graph_Reference_Condition
    print(paste("Graph_Reference_Condition is ", Graph_Reference_Condition,
                sep = ""))
  }

  # Set the Graph_Manual_Colour variable to either TRUE or FALSE
  # Defaults to FALSE

  if (missing(Graph_Manual_Colour)
      | is.null(Graph_Manual_Colour)) {
    Graph_Manual_Colour <- FALSE
    print(paste("No Graph_Manual_Colour input, ",
                "defaulting to Graph_Manual_Colour = FALSE",
                sep = ""))
  } else if ((Graph_Manual_Colour == TRUE)
             | (Graph_Manual_Colour == FALSE)) {
    Graph_Manual_Colour <- Graph_Manual_Colour
  } else {
    warning("Graph_Manual_Colour must be either TRUE of FALSE")
  }

  # If Graph_Manual_Colour is TRUE, it will override any Graph_Colour input
  # Warn the user if they have chosen a Graph_Colour as well as activating...
  # Graph_Manual_Colour

  if ((Graph_Manual_Colour == TRUE)
      & (!is.null(Graph_Colour))
      & (!missing(Graph_Colour))) {
    warning("Graph_Manual_Colour = TRUE overrides any Graph_Colour input")
  }

  # Graph_Manual_Colour is incompatible with Graph_Type == "Single"

  if ((Graph_Type == "Single")
      & (Graph_Manual_Colour == TRUE)) {
    stop("Graph_Manual_Colour is incompatible with Graph_Type == Single")
  }

  # There are multiple options for the Graph_Colour variable:

  Graph_Colour_options <- c("Viridis", "Brewer", "Grey")

  # Set the Graph_Colour variable to one of the above options
  # Defaults to "Viridis"
  # Graph_Colour must be one of these

  if (missing(Graph_Colour)
      | is.null(Graph_Colour)) {
    print("No Graph_Colour input, defaulting to Graph_Colour = 'Viridis'")
    Graph_Colour <- "Viridis"
  } else if (Graph_Colour %in% Graph_Colour_options) {
    Graph_Colour <- Graph_Colour
  } else {
    stop("Graph_Colour must be one of: ", paste0(Graph_Colour_options,
                                                 collapse = ", "))
  }

  # Set the Graph_Palette variable
  # There are 4 Graph_Colour options
  # Brewer and Viridis can incorporate many colour palettes
  # These are chosen by Graph_Palette
  # See individual colour schemes for details

  if ((missing(Graph_Palette))
      | is.null(Graph_Palette)) {
    Graph_Palette <- "viridis"
    print(paste("No Graph_Palette_Begin input, ",
                "defaulting to Graph_Palette = 'viridis'",
                sep = ""))
    if (Graph_Colour == "Brewer") {
      Graph_Palette <- "Greens"
      print(paste("The default Graph_Palette 'viridis' is not applicable with",
                  " Graph_Colour = Brewer. The Brewer palette 'Greens' will",
                  " be chosen", sep = ""))
    }
  }

  # Set the Graph_Palette_Begin variable
  # Defaults to 0
  # Must be numeric and between 0 and 1

  if (missing(Graph_Palette_Begin)
      | is.null(Graph_Palette_Begin)) {
    Graph_Palette_Begin <- 0
    print("No Graph_Palette_Begin input, defaulting to Graph_Palette_Begin = 1")
  } else if (is.numeric(Graph_Palette_Begin)
             & (Graph_Palette_Begin >= 0)
             & (Graph_Palette_Begin <= 1)) {
    Graph_Palette_Begin <- Graph_Palette_Begin
  } else {
    stop("Graph_Palette_Begin must be a numeric value between 0 and 1")
  }

  # Set the Graph_Palette_End variable
  # Defaults to 1
  # Must be numeric and between 0 and 1

  if (missing(Graph_Palette_End)
      | is.null(Graph_Palette_End)) {
    Graph_Palette_End <- 1
    print("No Graph_Palette_End input, defaulting to Graph_Palette_End = 1")
  } else if (is.numeric(Graph_Palette_End)
             & (Graph_Palette_End >= 0)
             & (Graph_Palette_End <= 1)) {
    Graph_Palette_End <- Graph_Palette_End
  } else {
    stop("Graph_Palette_End must be a numeric value between 0 and 1")
  }

  # Set the Graph_Text_Colour variable
  # Any hexcode or color() option are suitable, it is not feasible to check...
  # the validity of this user input. The warning given if invalid is...
  # acceptable.
  # Defaults to "grey30"

  if (missing(Graph_Text_Colour)
      | is.null(Graph_Text_Colour)) {
    Graph_Text_Colour <- "grey30"
    print(paste("No Graph_Text_Colour input, ",
                "defaulting to Graph_Text_Colour = 'grey30'",
                sep = ""))
  } else {
    Graph_Text_Colour <- Graph_Text_Colour
  }

  # Set the Graph_Background_Colour variable
  # Any hexcode or color() option are suitable, it is not feasible to check...
  # the validity of this user input. The warning given if invalid is...
  # acceptable.
  # Defaults to "grey95"

  if (missing(Graph_Background_Colour)
      | is.null(Graph_Background_Colour)) {
    Graph_Background_Colour <- "grey95"
    print(paste("No Graph_Background_Colour input, ",
                "defaulting to Graph_Background_Colour = 'grey95'",
                sep = ""))
  } else {
    Graph_Background_Colour <- Graph_Background_Colour
  }

  # Set the Graph_Hline_Colour variable
  # Any hexcode or color() option are suitable, it is not feasible to check...
  # the validity of this user input. The warning given if invalid is...
  # acceptable.
  # Defaults to "grey30"

  if (missing(Graph_Hline_Colour)
      | is.null(Graph_Hline_Colour)) {
    Graph_Hline_Colour <- "grey30"
    print(paste("No Graph_Hline_Colour input, ",
                "defaulting to Graph_Hline_Colour = 'grey30'",
                sep = ""))
  } else {
    Graph_Hline_Colour <- Graph_Hline_Colour
  }

  # Set the Graph_Legend variable to either TRUE or FALSE
  # Defaults to TRUE

  if (missing(Graph_Legend)
      | is.null(Graph_Legend)) {
    print("No Graph_Legend input, defaulting to Graph_Legend = TRUE")
    Graph_Legend <- TRUE
  } else if ((Graph_Legend == TRUE)
             | (Graph_Legend == FALSE)) {
    Graph_Legend <- Graph_Legend
  } else {
    stop("Graph_Legend must be either TRUE or FALSE")
  }

  # Set the Graph_Size_Right_Label variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Right_Label must be numeric and positive

  if (missing(Graph_Size_Right_Label)
      | is.null(Graph_Size_Right_Label)) {
    Graph_Size_Right_Label <- 1
    print(paste("No Graph_Size_Right_Label input, ",
                "defaulting to Graph_Size_Right_Label = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Right_Label)
             & (Graph_Size_Right_Label > 0)) {
    Graph_Size_Right_Label <- Graph_Size_Right_Label
  } else {
    stop("Graph_Size_Right_Label must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Top_Label variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Top_Label must be numeric and positive

  if (missing(Graph_Size_Top_Label)
      | is.null(Graph_Size_Top_Label)) {
    Graph_Size_Top_Label <- 1
    print(paste("No Graph_Size_Top_Label input, ",
                "defaulting to Graph_Size_Top_Label = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Top_Label)
             & (Graph_Size_Top_Label > 0)) {
    Graph_Size_Top_Label <- Graph_Size_Top_Label
  } else {
    stop("Graph_Size_Top_Label must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Y_Axis variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Y_Axis must be numeric and positive

  if (missing(Graph_Size_Y_Axis)
      | is.null(Graph_Size_Y_Axis)) {
    Graph_Size_Y_Axis <- 1
    print(paste("No Graph_Size_Y_Axis input, ",
                "defaulting to Graph_Size_Y_Axis = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Y_Axis)
             & (Graph_Size_Y_Axis > 0)) {
    Graph_Size_Y_Axis <- Graph_Size_Y_Axis
  } else {
    stop("Graph_Size_Y_Axis must be a numeric value greater than 0")
  }

  # Set the Graph_Size_X_Axis variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_X_Axis must be numeric and positive

  if (missing(Graph_Size_X_Axis)
      | is.null(Graph_Size_X_Axis)) {
    Graph_Size_X_Axis <- 1
    print(paste("No Graph_Size_X_Axis input, ",
                "defaulting to Graph_Size_X_Axis = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_X_Axis)
             & (Graph_Size_X_Axis > 0)) {
    Graph_Size_X_Axis <- Graph_Size_X_Axis
  } else {
    stop("Graph_Size_X_Axis must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Legend variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Legend must be numeric and positive

  if (missing(Graph_Size_Legend)
      | is.null(Graph_Size_Legend)) {
    Graph_Size_Legend <- 1
    print(paste("No Graph_Size_Legend input, ",
                "defaulting to Graph_Size_Legend = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Legend)
             & (Graph_Size_Legend > 0)) {
    Graph_Size_Legend <- Graph_Size_Legend
  } else {
    stop("Graph_Size_Legend must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Legend_Text variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Legend_Text must be numeric and positive

  if (missing(Graph_Size_Legend_Text)
      | is.null(Graph_Size_Legend_Text)) {
    Graph_Size_Legend_Text <- 1
    print(paste("No Graph_Size_Legend_Text input, ",
                "defaulting to Graph_Size_Legend_Text = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Legend_Text)
             & (Graph_Size_Legend_Text > 0)) {
    Graph_Size_Legend_Text <- Graph_Size_Legend_Text
  } else {
    stop("Graph_Size_Legend_Text must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Percentages variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Percentages must be numeric and positive

  if (missing(Graph_Size_Percentages)
      | is.null(Graph_Size_Percentages)) {
    Graph_Size_Percentages <- 1
    print(paste("No Graph_Size_Percentages input, ",
                "defaulting to Graph_Size_Percentages = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Percentages)
             & (Graph_Size_Percentages > 0)) {
    Graph_Size_Percentages <- Graph_Size_Percentages
  } else {
    stop("Graph_Size_Percentages must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Statistics variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Statistics must be numeric and positive

  if (missing(Graph_Size_Statistics)
      | is.null(Graph_Size_Statistics)) {
    Graph_Size_Statistics <- 1
    print(paste("No Graph_Size_Statistics input, ",
                "defaulting to Graph_Size_Statistics = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Statistics)
             & (Graph_Size_Statistics > 0)) {
    Graph_Size_Statistics <- Graph_Size_Statistics
  } else {
    stop("Graph_Size_Statistics must be a numeric value greater than 0")
  }

  # Set the Graph_Size_Datapoints variable, a multiplier.
  # Defaults to 1.
  # Graph_Size_Datapoints must be numeric and positive

  if (missing(Graph_Size_Datapoints)
      | is.null(Graph_Size_Datapoints)) {
    Graph_Size_Datapoints <- 1
    print(paste("No Graph_Size_Datapoints input, ",
                "defaulting to Graph_Size_Datapoints = 1",
                sep = ""))
  } else if (is.numeric(Graph_Size_Datapoints)
             & (Graph_Size_Datapoints > 0)) {
    Graph_Size_Datapoints <- Graph_Size_Datapoints
  } else {
    stop("Graph_Size_Datapoints must be a numeric value greater than 0")
  }

  # Set the Graph_Output variable to either TRUE or FALSE
  # Defaults to FALSE

  if (missing(Graph_Output)
      | is.null(Graph_Output)) {
    print("No Graph_Output input, defaulting to Graph_Output = FALSE")
    Graph_Output <- FALSE
  } else if ((Graph_Output == TRUE)
             | (Graph_Output == FALSE)) {
    Graph_Output <- Graph_Output
  } else {
    stop("Graph_Output must be either TRUE or FALSE")
  }

  # Set the Graph_Resolution variable
  # Defaults to 1200 for a high quality image
  # Graph_Resolution must be numeric and positive

  if (missing(Graph_Resolution)
      | is.null(Graph_Resolution)) {
    Graph_Resolution <- 1200
    print("No Graph_Resolution input, defaulting to Graph_Resolution = '1200'")
  } else if (is.numeric(Graph_Resolution)
             & (Graph_Resolution > 0)) {
    Graph_Resolution <- Graph_Resolution
  } else {
    stop("Graph_Resolution must be a numeric value greater than 0")
  }

  # Set the Graph_Width_Adjustment variable
  # Defaults to 1
  # Graph_Width_Adjustment must be numeric and positive

  if (missing(Graph_Width_Adjustment)
      | is.null(Graph_Width_Adjustment)) {
    Graph_Width_Adjustment <- 1
    print(paste("No Graph_Width_Adjustment input, ",
                "defaulting to Graph_Width_Adjustment = 1",
                sep = ""))
  } else if (is.numeric(Graph_Width_Adjustment)
             & (Graph_Width_Adjustment > 0)) {
    Graph_Width_Adjustment <- Graph_Width_Adjustment
  } else {
    stop("Graph_Width_Adjustment must be a numeric value greater than 0")
  }

  # Set the Graph_Height_Adjustment variable
  # Defaults to 1
  # Graph_Height_Adjustment must be numeric and positive

  if (missing(Graph_Height_Adjustment)
      | is.null(Graph_Height_Adjustment)) {
    Graph_Height_Adjustment <- 1
    print(paste("No Graph_Height_Adjustment input, ",
                "defaulting to Graph_Height_Adjustment = 1",
                sep = ""))
  } else if (is.numeric(Graph_Height_Adjustment)
             & (Graph_Height_Adjustment > 0)) {
    Graph_Height_Adjustment <- Graph_Height_Adjustment
  } else {
    stop("Graph_Height_Adjustment must be a numeric value greater than 0")
  }

  # Set the Graph_File variable
  # Defaults to "Colonisation Percentage Graph {Experiment_Name}.png"

  if (missing(Graph_File)
      | is.null(Graph_File)) {
    print(paste0("No Graph_File input, defaulting to ",
                 "Colonisation Percentage Graph {Experiment_Name}.png"))
    Graph_File <- glue("Colonisation Percentage Graph {Experiment_Name}.png")
    Graph_File <- gsub(x = Graph_File, " .png", ".png")
  } else {
    Graph_File <- paste(Graph_File, sep = "")
  }

  #####
  # B) Create a sub-function called Pull_Statistical_Information
  #####

  # This function:
  # Takes the output of TukeyHSD, Dunn and Wilcoxon tests performed later and...
  # creates a new output with the statistical letter groups and * for...
  # each condition. This function takes 5 inputs, the Statistical_Test which...
  # produced the data to analyse, the Reference_Stat_Group,...
  # the dependent variable, the data to analyse and the output name.
  # This data builds into the "OUT" excel spreadsheet that is established
  # later in the AMReader function and creates a new dataframe

  Pull_Statistical_Information <- function(
    Statistical_Test,
    Reference_Stat_Group = NULL,
    Variable,
    Statistics_Dataset,
    Output_name
  )

  {

    # Set certain variables depending on the statistical test used to produce...
    # the data being analysed.
    # These identify the column names required to process data from the...
    # statistical test, and identifies the rows for conditional formatting...
    # in the excel output

    if (Statistical_Test == "Tukey") {
      p.value_name <- "p.adj"
      comparison_name <- "comparisons"
      Format_Columns <- 5
      Format_Rows <- 1:1000

    } else if (Statistical_Test == "Dunn") {
      p.value_name <- "pval"
      comparison_name <- "comparison"
      Format_Columns <- 3
      Format_Rows <- 1:1000

    } else if (Statistical_Test == "Wilcoxon") {
      Format_Columns <- 1:100
      Format_Rows <- 1:100
    }

    # Prepare the name of the new sheet based on the Variable being processed

    Tab_name <- paste(Statistical_Test, "_", Variable, sep = "")

    # Create a new sheet in excel according to the Tab_name

    addWorksheet(OUT, Tab_name)

    # Write the Stat_Test results to excel

    if ((Statistical_Test == "Tukey")
        | (Statistical_Test == "Dunn")) {

      writeData(OUT, sheet = Tab_name, x = Statistics_Dataset, rowNames = FALSE)

    } else if (Statistical_Test == "Wilcoxon") {

      writeData(OUT, sheet = Tab_name, x = Statistics_Dataset, rowNames = TRUE)
    }

    # Highlight p values which are p < 0.05

    conditionalFormatting(OUT,
                          cols = Format_Columns,
                          rows = Format_Rows,
                          sheet = Tab_name,
                          type = "expression",
                          rule = "<0.05",
                          style = Statistically_significant)
    conditionalFormatting(OUT,
                          cols = Format_Columns,
                          rows = Format_Rows,
                          sheet = Tab_name,
                          type = "expression",
                          rule = "==0",
                          style = Empty)
    conditionalFormatting(OUT,
                          cols = Format_Columns,
                          rows = Format_Rows,
                          sheet = Tab_name,
                          type = "expression",
                          rule = ">=0.05",
                          style = Borders)

    # Prepare a unique column name based on the variable being processed

    Letter_Column_name <- paste(Statistical_Test, "group", Variable, sep = "_")

    Star_Column_name <- paste(Statistical_Test, "star_group", Variable, sep = "_")

    # Compute significance groups and tidy the resulting dataframe

    if ((Statistical_Test == "Tukey")
        | (Statistical_Test == "Dunn")) {

      Significance_Groups <- cldList(
        p.value = Statistics_Dataset[, p.value_name],
        comparison = Statistics_Dataset[ , comparison_name],
        data = Statistics_Dataset,
        threshold = 0.05,
        remove.space = FALSE,
        remove.zero = FALSE) %>%
        rename(Stat_Test_Friendly = Group) %>%
        rename_with(.fn = ~paste(
          Letter_Column_name,
          sep = ""),
          .cols = "Letter") %>%
        select(all_of(c("Stat_Test_Friendly", Letter_Column_name)))

    }

    if (Statistical_Test == "Wilcoxon") {

      # Convert Wilcoxon results into a format suitable for...
      # forming letter groups

      Full_Matrix <- fullPTable(Wilcoxon_Test)

      # Compute significance groups and tidy the resulting dataframe

      Significance_Groups <- as.data.frame(
        multcompLetters(Full_Matrix, threshold = 0.05)["Letters"]) %>%
        rownames_to_column("Stat_Test_Friendly") %>%
        rename_with(.fn = ~paste(
          "Wilcoxon_group_",
          Variable,
          sep = ""),
          .cols = "Letters")

    }

    # If Reference_Stat_Group is applied:

    if (!is.null(Reference_Stat_Group)) {

      # Highlight cells in the excel output which contain that reference...
      # condition to aid visualisation

      conditionalFormatting(OUT,
                            cols = 1:100,
                            rows = 1:1000,
                            sheet = Tab_name,
                            type = "contains",
                            rule = Reference_Stat_Group,
                            style = Reference)

      # Determine the letter group(s) of the Reference_Stat_Group

      Graph_Reference_Condition_Letter <- Significance_Groups %>%
        filter(Stat_Test_Friendly == Reference_Stat_Group) %>%
        select(all_of(Letter_Column_name)) %>%
        unique() %>%
        pull()

      # Create a new column to list * if there is a statistical difference...
      # between the given condition and the reference condition. This column
      # is named after the Star_Column_name

      # It is necessary to turn the Graph_Reference_Condition_Letter
      # into a character of form "a|b|c" or similar, so as to test whether...
      # there is any overlap with the query Letter group

      Graph_Reference_Condition_Letter <- paste(
        as.list(
          strsplit(
            Graph_Reference_Condition_Letter, ""))[[1]],
        collapse = "|")

      # If any of the letters in Graph_Reference_Condition_Letter match the...
      # query, provide "", else provide *

      Significance_Groups <- Significance_Groups %>%
        mutate(Stars = case_when(
          grepl(Graph_Reference_Condition_Letter,
                Significance_Groups[ , Letter_Column_name]) ~ "",
          TRUE ~ "*")) %>%
        rename_with(.fn = ~paste(Star_Column_name), .cols = Stars)

    }

    assign(Output_name,
           Significance_Groups, envir= parent.frame())

  }

  #####
  # C) Load data and store data in Global Environment
  #####

  # Load colonisation count data based off the given path from Chosen_Path

  Count_data <- read_excel(
    path  = Chosen_Path,
    sheet = "AM Results"
  )

  # Load metadata about the conditions from the same excel document

  Condition_data <- read_excel(
    path  = Chosen_Path,
    sheet = "Conditions"
  )

  # Save data to global environment for user

  assign("Count_data", Count_data, envir=.GlobalEnv)
  assign("Condition_data", Condition_data, envir=.GlobalEnv)

  #####
  # D) Process information about the Facets from the user
  #####

  # Create a list of all possible Facets

  Facet_List <- list("Facet_1", "Facet_2", "Facet_3")

  # Create a list of Facets from the user input

  Facet_Input <- list(Facet_1, Facet_2, Facet_3)

  # Create a list of Facets to include from this information

  Facets_Include <- do.call(rbind,
                            Map(data.frame,
                                Facet = Facet_List,
                                Input = Facet_Input)) %>%
    subset(Input == TRUE) %>%
    pull(Facet)

  # Check that the input of Facet_X_Order is contained within the data,...
  # where X is 1, 2 or 3
  # If there are invalid values, provide an error message to the user

  invalid_facets_1 <- Graph_Facet_1_Order[!(Graph_Facet_1_Order %in%
                                              Condition_data$Facet_1)]

  if (length(invalid_facets_1) > 0) {
    stop(paste(
      "The following values from Graph_Facet_1_Order",
      " do not appear in the data: ",
      paste0(invalid_facets_1, collapse = ", "),
      " - Please provide values which appear in Condition_data",
      sep = ""))
  }

  invalid_facets_2 <- Graph_Facet_2_Order[!(Graph_Facet_2_Order %in%
                                              Condition_data$Facet_2)]

  if (length(invalid_facets_2) > 0) {
    stop(paste(
      "The following values from Graph_Facet_2_Order",
      " do not appear in the data: ",
      paste0(invalid_facets_2, collapse = ", "),
      " - Please provide values which appear in Condition_data",
      sep = ""))
  }

  invalid_facets_3 <- Graph_Facet_3_Order[!(Graph_Facet_3_Order %in%
                                              Condition_data$Facet_3)]

  if (length(invalid_facets_3) > 0) {
    stop(paste(
      "The following values from Graph_Facet_3_Order",
      " do not appear in the data: ",
      paste0(invalid_facets_3, collapse = ", "),
      " - Please provide values which appear in Condition_data",
      sep = ""))
  }

  #####
  # E) Process the data to generate the Filtered_Dataset
  #####

  # Remove rows of data for slides that have not been counted
  # Ensure rows with data are numeric - this enables backwards compatability
  # with an earlier version of AMScorer

  Count_data_included <- Count_data %>%
    subset(Counted == 1) %>%
    rename("Letter" = "Condition") %>%
    mutate(Total = as.numeric(Total)) %>%
    mutate(EH = as.numeric(EH)) %>%
    mutate(H = as.numeric(H)) %>%
    mutate(IH = as.numeric(IH)) %>%
    mutate(A = as.numeric(A)) %>%
    mutate(V = as.numeric(V)) %>%
    mutate(S = as.numeric(S))

  # Calculate number of replicates per condition

  Count_data_samplesize <- Count_data_included %>%
    group_by(Letter) %>%
    summarise(Length = length(Total))

  # Attach metadata information to Count_data_samplesize and establish a new...
  # piece of metadata with condition and number of replicates which could...
  # act as a label in the final graph if Graph_Sample_Sizes == TRUE
  # Condition_Size is of form (WT n = 6)

  metadata <- left_join(
    Count_data_samplesize,
    Condition_data,
    by = "Letter") %>%
    mutate(Condition_Size = paste(Condition,
                                  paste0("(n = ", Length, ")"),
                                  sep = " "))

  # Based of user input, create the graph Label including either the...
  # condition alone or condition with the number of replicates

  if (Graph_Sample_Sizes == TRUE) {
    metadata <- metadata %>%
      mutate(Label = Condition_Size)
  } else {
    metadata <- metadata %>%
      mutate(Label = Condition)
  }

  # Join count data with metadata
  # Replace Condition with the defined Label (which may or may not include...
  # Graph_Sample_Sizes)
  # Save the user input of Condition as Condition_original

  Counts_and_metadata <- left_join(
    Count_data_included, metadata, by = "Letter") %>%
    rename(Condition_original = Condition) %>%
    rename(Condition = Label)

  # The statistics computed by this function will attempt to separate each...
  # slide of data into an individual group. It will do this based off the...
  # Condition supplied, as well as the Facets. It will only include the...
  # facets given by the user's input - if no facets are included, it will...
  # group all data with the same "Condition" name into one group. This group...
  # is then called "Statistcial_Group" and is formed by collapsing the...
  # text in Condition and the included Facets.

  Counts_and_metadata_stat_group <- Counts_and_metadata %>%
    mutate(Facet_Collapse = apply(
      Counts_and_metadata[ , Facets_Include],
      1,
      paste,
      collapse = "-")) %>%
    mutate(Statistical_Group = paste(Condition_original, Facet_Collapse))

  # If Condition_Exclude has been specified, use it to establish...
  # Conditions_to_include

  if ((!missing(Condition_Exclude))
      & (!is.null(Condition_Exclude))) {
    Conditions_to_include <- unlist(list(
      Counts_and_metadata_stat_group %>%
        subset(!(Letter %in% Condition_Exclude)) %>%
        select(Letter) %>%
        unique()
    ))
  }

  # Select only the conditions of interest, as defined by the user

  Filtered_Dataset <- Counts_and_metadata_stat_group %>%
    subset(Letter %in% Conditions_to_include)

  # Create a tidier dataframe with desirable columns and metadata

  Filtered_Dataset <- Filtered_Dataset %>%
    select(all_of(c("Letter", "Slide", "Condition_original", "Condition",
                    Structures_list,
                    "Facet_1", "Facet_2", "Facet_3", "Manual_Colour",
                    "Facet_Collapse", "Statistical_Group")))

  # The downstream statistical analyses employs cldList and multcompLetters...
  # which cannot operate on names which include '-'. To side-step this...
  # all '-' will be converted into '_' in a new column called...
  # Stat_Test_Friendly.
  # The Stat_Test_Friendly column is converted into factors.

  Filtered_Dataset <- Filtered_Dataset %>%
    mutate(Stat_Test_Friendly = gsub('-', '_', Statistical_Group)) %>%
    mutate(Stat_Test_Friendly = as.factor(Stat_Test_Friendly))

  # Save data to global environment for user

  assign("Filtered_Dataset",
         Filtered_Dataset, envir=.GlobalEnv)

  # Test if Facet_1, Facet_2 or Facet_3 should be active
  # If metadata has been provided in Conditions_data for Facets...
  # Give warnings to suggest

  Facet_1_Data <- unlist(list(
    Filtered_Dataset %>%
      select(Facet_1) %>%
      na.omit() %>%
      unique()
  ))

  Facet_2_Data <- unlist(list(
    Filtered_Dataset %>%
      select(Facet_2) %>%
      na.omit() %>%
      unique()
  ))

  Facet_3_Data <- unlist(list(
    Filtered_Dataset %>%
      select(Facet_3) %>%
      na.omit() %>%
      unique()
  ))

  if ((Facet_1 == TRUE)
      & (length(Facet_1_Data) <= 1)) {
    warning(paste("Facet_1 = TRUE, but no more than one unique variable",
                  " has been found for Facet_1.",
                  " This would suggest that Facet_1 should be FALSE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  } else if ((Facet_1 == FALSE)
             & (length(Facet_1_Data) > 1)) {
    warning(paste("Facet_1 = FALSE, but more than one unique variable",
                  " has been found for Facet_1.",
                  " This would suggest that Facet_1 should be TRUE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  }

  if ((Facet_2 == TRUE)
      & (length(Facet_2_Data) <= 1)) {
    warning(paste("Facet_2 = TRUE, but no more than one unique variable",
                  " has been found for Facet_2.",
                  " This would suggest that Facet_2 should be FALSE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  } else if ((Facet_2 == FALSE)
             & (length(Facet_2_Data) > 1)) {
    warning(paste("Facet_2 = FALSE, but more than one unique variable",
                  " has been found for Facet_2.",
                  " This would suggest that Facet_2 should be TRUE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  }

  if ((Facet_3 == TRUE)
      & (length(Facet_3_Data) <= 1)) {
    warning(paste("Facet_3 = TRUE, but no more than one unique variable",
                  " has been found for Facet_3.",
                  " This would suggest that Facet_3 should be FALSE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  } else if ((Facet_3 == FALSE)
             & (length(Facet_3_Data) > 1)) {
    warning(paste("Facet_3 = FALSE, but more than one unique variable",
                  " has been found for Facet_3.",
                  " This would suggest that Facet_3 should be TRUE.",
                  " Please check this and amend if necessary.",
                  sep = ""))
  }

  #####
  # F) Create the Condition_Order_list
  #####

  # To generate the plot, a list of Condition labels that describe...
  # the order of Conditions desired by the user is needed

  Condition_Order_list <- rbind(
    metadata %>%
      subset(Letter %in% Graph_Condition_Order) %>%
      unique() %>%
      arrange(factor(Letter, levels = Graph_Condition_Order)),
    metadata %>%
      subset(!(Letter %in% Graph_Condition_Order)) %>%
      unique() %>%
      arrange(factor(Letter, levels = Graph_Condition_Order))) %>%
    pull(Label)

  # The Condition_Order_list must contain unique values

  Non_Duplicated_Condition_Order_list <-
    Condition_Order_list[!duplicated(Condition_Order_list)]

  #####
  # G) Statistical Analysis
  #####

  # Create a new dataframe to collect results

  Statistics_Results <- Filtered_Dataset

  # Establish the Stat_Test_Friendly value which is the...
  # Graph_Reference_Condition. This is called Reference_Stat_Group.
  # It defaults to NULL

  # Check the Graph_Reference_Condition specified by the user is found...
  # in the data

  if (!is.null(Graph_Reference_Condition)) {

    if (Graph_Reference_Condition %in% Filtered_Dataset$Letter) {

      # Pull the Stat_Test_Friendly group

      Reference_Stat_Group <- as.character(
        Filtered_Dataset %>%
          subset(Letter == Graph_Reference_Condition) %>%
          select(Stat_Test_Friendly) %>%
          unique() %>%
          droplevels() %>%
          pull(Stat_Test_Friendly))

      # Save data to global environment for user

      assign("Reference_Stat_Group",
             Reference_Stat_Group, envir=.GlobalEnv)

      # If Graph_Reference_Condition does not appear in the data

    } else {

      stop(paste("The Graph_Reference_Condition is '",
                 Graph_Reference_Condition,
                 "' however this does not appear in the data.",
                 " Please provide a letter code (A-Z, AA-AZ, BA-BZ) ",
                 "that appears in Filtered_Dataset",
                 sep = ""))
    }

    # If Graph_Stat_Display = "Reference", and Facets are applied, warn the...
    # user that Graph_Stat_Display = "Refernce" is not largely compatible
    # with Facet_1/2/3 being active as all groups are being compared to one...
    # group in one facet

    if ((Graph_Stat_Display == "Reference")
        & (length(Facets_Include) > 0)) {

      warning(paste("Graph_Stat_Display = 'Reference' may be misleading if",
                    " Facet_1/2/3 are active. See manual for further",
                    " information.", sep = ""))

    }

  } else {

    Reference_Stat_Group <- NULL

  }

  # None of the statistical analyses can be performed if there is a single...
  # statistical group. If this is the case, no stats will be conducted...
  # and a warning issued to the user. The Graph_Stat_Test variable will...
  # also be overwritten.

  if (length(unique(Filtered_Dataset$Stat_Test_Friendly)) <= 1) {
    warning(paste0("Only a single statistical_group has been identified. ",
                   "No stats will be computed"))
    Graph_Stat_Test <- NULL
  } else {

    # This else {} encapsulates all statistical analyses

    # The downstream statistical analyses are dependent on variation between...
    # statistical groups for each given fungal structure. If this is not...
    # the case, the statistics will fail. Hence, stats will only be...
    # computed if there are two or more different values within each Structure
    # The valid structures are identified in Structures_list_for_stats

    Structures_list_for_stats <- Structures_list %>%
      map_dfr(.f = function(.each_structure) {
        data.frame(
          Structure = .each_structure,
          Sum = as.numeric(
            sum(!duplicated(Filtered_Dataset[ , .each_structure])))
        )
      }) %>%
      filter(Sum > 1) %>%
      pull(Structure)

    # Structures are then compared to the user input from Structures_Include...
    # or Structures_Exclude and only those valid Structures which the user...
    # wishes to include are processed.

    Structures_list_for_stats <-
      Structures_list_for_stats[Structures_list_for_stats %in% Structures]

    # All stats that are computed are saved to an excel spreadsheet - OUT
    # Open workbook creation

    OUT <- createWorkbook()

    # Set styles for conditional formatting rules to highlight statistically...
    # significant results.

    Statistically_significant <- createStyle(bgFill = "#0EDFFF",
                                             border = "TopBottomLeftRight")

    # If the cell is empty, it will meet the criteria of being statistically...
    # significant (< 0.05). Return it back to white after it is coloured green

    Empty <- createStyle(bgFill = "#FFFFFF",
                         border = "TopBottomLeftRight")

    # Colour removes borders. These should afterwards be replaced

    Borders <- createStyle(bgFill = "#FFFFFF",
                           border = "TopBottomLeftRight")

    # Highlight cells referring to the reference condition

    Reference <- createStyle(bgFill = "#F3B700",
                             border = "TopBottomLeftRight")

    # Create a new sheet in excel to summarise the statistically relevant...
    # inputs of the function

    addWorksheet(OUT, "Statistical_Input")

    # Add an authorship note

    Authorship_Note <- paste0(
      "AMScorer and AMReader",
      " were developed by Edwin Jarratt-Barnham")

    writeData(OUT,
              sheet = "Statistical_Input",
              x = Authorship_Note,
              rowNames = FALSE)

    # Create a data.frame of statistically relevant inputs of the function

    Statistical_Input <- data.frame(
      User_Input = c("Stat_Sided",
                     "Stat_Dunn_Padj",
                     "Stat_Wilcoxon_Padj",
                     "Structures",
                     "Conditions"),
      Values = c(Stat_Sided,
                 Stat_Dunn_Padj,
                 Stat_Wilcoxon_Padj,
                 paste0(Structures, collapse = ", "),
                 paste0(Conditions_to_include, collapse = ", ")))

    # Write the Statistical_Input results to excel

    writeData(OUT, sheet = "Statistical_Input",
              x = Statistical_Input,
              rowNames = FALSE,
              startRow = 3)

    # G.1) ANOVA

    # The ANOVA formula for all tests will be Structure ~ Stat_Test_Friendly

    # Perform an ANOVA for each valid Structure, and extract the results...
    # into a single dataframe

    ANOVA_Results <- Structures_list_for_stats %>%
      map_dfr(.f = function(.each_structure) {
        AOV <- tidy(aov(
          reformulate(termlabels = "Stat_Test_Friendly",
                      response = .each_structure),
          data = Filtered_Dataset))
        data.frame(
          Structure = .each_structure,
          df = AOV[1:length("Stat_Test_Friendly"),"df"],
          sumsq = AOV[1:length("Stat_Test_Friendly"),"sumsq"],
          meansq = AOV[1:length("Stat_Test_Friendly"),"meansq"],
          F_value = AOV[1:length("Stat_Test_Friendly"),"statistic"],
          p.value = AOV[1:length("Stat_Test_Friendly"),"p.value"]
        )
      })

    # Create a new sheet in excel for these ANOVA results

    addWorksheet(OUT, "ANOVA_Results")

    # Write the ANOVA results to excel

    writeData(OUT,
              sheet = "ANOVA_Results",
              x = ANOVA_Results,
              rowNames = FALSE)

    # Highlight green the p values which are p < 0.05

    conditionalFormatting(OUT,
                          cols = 6,
                          rows = 1:1000,
                          sheet = "ANOVA_Results",
                          type = "expression",
                          rule = "<0.05",
                          style = Statistically_significant)

    conditionalFormatting(OUT,
                          cols = 6,
                          rows = 1:1000,
                          sheet = "ANOVA_Results",
                          type = "expression",
                          rule = "==0",
                          style = Empty)

    conditionalFormatting(OUT,
                          cols = 6,
                          rows = 1:1000,
                          sheet = "ANOVA_Results",
                          type = "expression",
                          rule = ">=0.05",
                          style = Borders)

    # G.2) ANOVA_Diagnostics and Tukey_HSD

    # Create a positioning variable for ANOVA_Diagnostics plot

    Insert_Position <- 1

    # Create a new sheet in excel for the ANOVA diagnostic plots

    addWorksheet(OUT, "ANOVA_Diagnostics")

    # For each valid structure, carry out the ANOVA and post-hoc TukeyHSD

    for (Structure in Structures_list_for_stats) {

      # Re-run ANOVA as previous

      ANOVA_Individual <- aov(
        reformulate(termlabels = "Stat_Test_Friendly",
                    response = Structure),
        data = Filtered_Dataset)

      # Create diagnostic ANOVA Plots
      # As this is only necessary if Stat_Output = TRUE, and is...
      # more time-intensive, only run if Stat_Output = TRUE

      if (Stat_Output == TRUE) {

        # Create a 2x2 grid for the 4 diagnostic plots

        par(mfrow = c(2, 2))

        # Set margins

        par(mar = c(2, 2, 2, 2))

        # Plot results

        plot(ANOVA_Individual)

        # Add text identifying the Structure being analysed

        mtext(Structure, side = 3, line = -1, outer = TRUE, adj = 0.1)

        # Add plots to OUT

        insertPlot(OUT, "ANOVA_Diagnostics",
                   startRow = Insert_Position,
                   height = 6,
                   width =  9)

        # Adjust Insert_Position for the next set of diagnostic plots

        Insert_Position <- Insert_Position + 33

      }

      # Run TukeyHSD

      Tukey_Individual <- TukeyHSD(x = ANOVA_Individual,
                                   "Stat_Test_Friendly",
                                   conf.level=0.95)

      # Convert the TukeyHSD output into a dataframe

      Tukey_Dataframe <- as.data.frame(Tukey_Individual$Stat_Test_Friendly)

      # Pull_Statistical_Information employs multcompLetters (cldList)...
      # to calculate different groups
      # This function cannot work if p.adj = Na, which occurs if there...
      # are no differences between the two conditions...
      # e.g. all values are 0
      # Therefore, these Na values will be converted to p.adj = 1

      Tukey_Dataframe[is.na(Tukey_Dataframe)] <- 1

      # Extract out the rownames and call them "comparisons"

      Tukey_Dataframe <- cbind('comparisons' = rownames(Tukey_Dataframe),
                               data.frame(Tukey_Dataframe, row.names=NULL))

      # Apply Pull_Statistical_Information function

      Pull_Statistical_Information(
        Statistical_Test = "Tukey",
        Reference_Stat_Group = Reference_Stat_Group,
        Variable = Structure,
        Statistics_Dataset = Tukey_Dataframe,
        Output_name = "Tukey_Significance_Groups"
      )

      # Add the TukeyHSD results onto the Statistics_Results dataframe

      Statistics_Results <- left_join(Statistics_Results,
                                      Tukey_Significance_Groups,
                                      by = "Stat_Test_Friendly")

    }

    # Save the significant letter groups from TukeyHSD to the excel output

    # Create a new list with column names for each Structure processed in the...
    # TukeyHSD for loop

    TukeyHSD_Groups_list <- apply(expand.grid("Tukey_group",
                                              Structures_list_for_stats),
                                  MARGIN = 1,
                                  paste0,
                                  collapse = "_")

    # Select these columns from the Statistics_Results dataframe

    TukeyHSD_Letter_Groups <- Statistics_Results %>%
      select(all_of(c("Stat_Test_Friendly", TukeyHSD_Groups_list))) %>%
      unique()

    # Create a new sheet in excel for TukeyHSD_Letter_Groups

    addWorksheet(OUT, "TukeyHSD_Letter_Groups")

    # Write the TukeyHSD_Letter_Groups to excel

    writeData(OUT, sheet = "TukeyHSD_Letter_Groups",
              x = TukeyHSD_Letter_Groups,
              rowNames = FALSE)

    # G.3) Kruskal-Wallis

    # Perform a Kruskal-Wallis test for each valid Structure, and extract the...
    # results into a single dataframe

    Kruskal_results <- Structures_list_for_stats %>%
      map_dfr(.f = function(.each_structure) {
        data.frame(
          Structure = .each_structure,
          p.value   = kruskal.test(
            formula = reformulate("Stat_Test_Friendly", .each_structure),
            data    = Filtered_Dataset
          )$p.value
        )
      })

    # Create a new sheet in excel for Kruskal_Wallis Results

    addWorksheet(OUT, "Kruskal_Wallis")

    writeData(OUT, sheet = "Kruskal_Wallis",
              x = Kruskal_results,
              rowNames = FALSE)

    # Highlight p values which are p < 0.05

    conditionalFormatting(OUT,
                          cols = 2,
                          rows = 1:1000,
                          sheet = "Kruskal_Wallis",
                          type = "expression",
                          rule = "<0.05",
                          style = Statistically_significant)
    conditionalFormatting(OUT, cols = 2,
                          rows = 1:1000,
                          sheet = "Kruskal_Wallis",
                          type = "expression",
                          rule = "==0",
                          style = Empty)
    conditionalFormatting(OUT, cols = 2,
                          rows = 1:1000,
                          sheet = "Kruskal_Wallis",
                          type = "expression",
                          rule = ">=0.05",
                          style = Borders)

    # G.4) Dunn Tests

    # For each valid structure, carry out the DunnTest

    for (Structure in Structures_list_for_stats) {

      # Carry out a Dunn Test using user chosen p-adjustment method

      Dunn_test <- DunnTest(
        x = as.numeric(unlist(Filtered_Dataset[ , Structure])),
        g = as.factor(unlist(Filtered_Dataset[ , "Stat_Test_Friendly"])),
        method = Stat_Dunn_Padj,
        alternative = Stat_Sided)

      # Create the table for the Pull_Statistical_Information function

      Dunn_test <- as.data.frame(Dunn_test[[1]]) %>%
        rownames_to_column("comparison")

      # Apply Pull_Statistical_Information function

      Pull_Statistical_Information(
        Statistical_Test = "Dunn",
        Reference_Stat_Group = Reference_Stat_Group,
        Variable = Structure,
        Statistics_Dataset = Dunn_test,
        Output_name = "Dunn_Significance_Groups"
      )

      Statistics_Results <- left_join(Statistics_Results,
                                      Dunn_Significance_Groups,
                                      by = "Stat_Test_Friendly")

    }

    # Save the significant letter groups from Dunn Test to the excel output

    # Create a new list with column names for each Structure processed in the...
    # Dunn Test for loop

    Dunn_Groups_list <- apply(expand.grid("Dunn_group",
                                          Structures_list_for_stats),
                              MARGIN = 1,
                              paste0,
                              collapse = "_")

    # Select these columns from the Statistics_Results dataframe

    Dunn_Letter_Groups <- Statistics_Results %>%
      select(all_of(c("Stat_Test_Friendly", Dunn_Groups_list))) %>%
      unique()

    # Create a new sheet in excel for Dunn_Letter_Groups

    addWorksheet(OUT, "Dunn_Letter_Groups")

    # Write the Dunn_Letter_Groups to excel

    writeData(OUT,
              sheet = "Dunn_Letter_Groups",
              x = Dunn_Letter_Groups,
              rowNames = FALSE)

    # G.5) Pairwise Wilcoxon Tests

    # For each valid structure, carry out the pairwise.wilcox.test

    for (Structure in Structures_list_for_stats) {

      # Carry out a Wilcoxon Test using user chosen p-adjustment method

      Wilcoxon_Test <- pairwise.wilcox.test(
        Filtered_Dataset[[Structure]],
        Filtered_Dataset[["Stat_Test_Friendly"]],
        p.adjust.method = Stat_Wilcoxon_Padj,
        alternative = Stat_Sided,
        exact = FALSE)$p.value

      # If all values are tied between two statistical groups e.g. all values...
      # are 0, pairwise.wilcoxon.test cannot compute a p.value, giving...
      # NaN. All NaN are therefore converted into 1 to allow...
      # multcompLetters to function

      Wilcoxon_Test[is.nan(Wilcoxon_Test)] <- 1

      # Create a dataframe to export to excel

      Wilcoxon_Dataframe <- as.data.frame(Wilcoxon_Test)

      # Apply Pull_Statistical_Information function

      Pull_Statistical_Information(
        Statistical_Test = "Wilcoxon",
        Reference_Stat_Group = Reference_Stat_Group,
        Variable = Structure,
        Statistics_Dataset = Wilcoxon_Dataframe,
        Output_name = "Wilcoxon_Significance_Groups"
      )

      # Add Wilcoxon results onto the Statistics_Results dataframe

      Statistics_Results <- left_join(Statistics_Results,
                                      Wilcoxon_Significance_Groups,
                                      by = "Stat_Test_Friendly")

    }

    # Save the significant letter groups from Wilcoxon tests to the excel output

    # Create a new list with column names for each Structure processed in the...
    # Wilcoxon Test for loop

    Wilcoxon_Groups_list <- apply(expand.grid("Wilcoxon_group",
                                              Structures_list_for_stats),
                                  MARGIN = 1,
                                  paste0,
                                  collapse = "_")

    # Select these columns from the Statistics_Results dataframe

    Wilcoxon_Letter_Groups <- Statistics_Results %>%
      select(all_of(c("Stat_Test_Friendly", Wilcoxon_Groups_list))) %>%
      unique()

    # Create a new sheet in excel for Wilcoxon_Letter_Groups

    addWorksheet(OUT, "Wilcoxon_Letter_Groups")

    # Write the Wilcoxon_Letter_Groups to excel

    writeData(OUT,
              sheet = "Wilcoxon_Letter_Groups",
              x = Wilcoxon_Letter_Groups,
              rowNames = FALSE)

    # G.6) Save the statistics data

    # Create the name for the excel output file

    Stat_Output_name <- glue("{Path}/{Stat_File}")

    # Save the results of the stats tests

    if (Stat_Output == TRUE) {
      saveWorkbook(OUT, Stat_Output_name, overwrite = TRUE)
    }

    # The statistical analyses above were dependent on variation between...
    # statistical groups for each given fungal structure and only...
    # calculated if there was variation. It is still desired to plot these...
    # data, which requires ggplot to receive data concerning all letter...
    # groups. An empty cell will be created for each case where statistics...
    # could not be calculated.

    # Create a list of structures for which stats have not been calculated

    Missing_Structures <- Structures_list[!(Structures_list %in%
                                              Structures_list_for_stats)]

    # Add a column for each statistical test into Statistics_Results with an...
    # empty cell, and do the same for star_group

    for (Structure in Missing_Structures) {

      Tukey_name <- paste("Tukey_group_", Structure, sep = "")
      Dunn_name <- paste("Dunn_group_", Structure, sep = "")
      Wilcoxon_name <- paste("Wilcoxon_group_", Structure, sep = "")
      Statistics_Results[[Tukey_name]] = ""
      Statistics_Results[[Dunn_name]] = ""
      Statistics_Results[[Wilcoxon_name]] = ""

      Tukey_name_star <- paste("Tukey_star_group_", Structure, sep = "")
      Dunn_name_star <- paste("Dunn_star_group_", Structure, sep = "")
      Wilcoxon_name_star <- paste("Wilcoxon_star_group_", Structure, sep = "")
      Statistics_Results[[Tukey_name_star]] = ""
      Statistics_Results[[Dunn_name_star]] = ""
      Statistics_Results[[Wilcoxon_name_star]] = ""

    }

    # Save data to global environment for user

    assign("Statistics_Results",
           Statistics_Results, envir=.GlobalEnv)

  }

  #####
  # H) Prepare data for Plot
  #####

  # Create a dataset for the ggplot containing the necessary columns
  # Pivot according to Structures
  # Re-order factors of Condition according to the Non_Duplicated_Order_list...
  # as determined by the user
  # Re-order the factors of Structures to a sensible order
  # Replace the letter descriptions of the Structures with their full name

  Plot_Dataset <- Filtered_Dataset %>%
    select(all_of(c("Condition",
                    Structures,
                    "Facet_1",
                    "Facet_2",
                    "Facet_3",
                    "Manual_Colour",
                    "Stat_Test_Friendly"))) %>%
    pivot_longer(cols = all_of(Structures),
                 names_to = "variable") %>%
    mutate(Condition = factor(Condition,
                              levels = Non_Duplicated_Condition_Order_list)) %>%
    mutate(variable = factor(variable,
                             levels = c("Total",
                                        "EH",
                                        "H",
                                        "IH",
                                        "A",
                                        "V",
                                        "S"))) %>%
    mutate(variable = revalue(
      variable,
      c("Total" = "Total",
        "EH"    = "Extraradical hyphae",
        "H"     = "Hyphopodia",
        "IH"    = "Intraradical hyphae",
        "A"     = "Arbuscules",
        "V"     = "Vesicles",
        "S"     = "Spores"
      )
    ))

  # It is necessary for there to be sufficient space on the graph to include...
  # statistical information. If the data values are very high, it is...
  # necessary to increase the y-axis beyond 100% to create room.

  # Determine the highest value in the data

  Max_value <- Filtered_Dataset %>%
    select(all_of(Structures)) %>%
    max()

  # If displaying statistical information on the graph - only when...
  # Graph_Stat_Test is not NULL and Graph_Type = "Facets" - Then assign...
  # the graph y-axis length and position for the statistical information...
  # accordingly. Otherwise, set the y-axis length based on max value. Adjust
  # the hlines to display depending on the range of the y-axis.

  if ((!is.null(Graph_Stat_Test))
      & (Max_value > 90)
      & (Graph_Type == "Facets")) {

    Stat_position  <- 105
    Coordinates <- c(0,110)
    Hlines <- c(25,50,75,100)

  } else if (Max_value > 95) {

    Coordinates <- c(0,110)
    Hlines <- c(25,50,75,100)

  } else {

    Stat_position <- 95
    Coordinates <- c(0,100)
    Hlines <- c(25,50,75)

  }

  # If Graph_Stat_Test refers to one of the statistical tests, then it is...
  # necessary to create a dataframe containing this information for ggplot.

  if (!is.null(Graph_Stat_Test)) {

    # The statistics will either be letter groups, or asterisks.
    # Create a variable Letters_or_Group to pull columns from Statistics_Results

    if (Graph_Stat_Display == "Letters") {

      # Pull columns for letters

      Letters_or_Group <- "group"

    } else if (Graph_Stat_Display == "Reference") {

      # Pull columns for asterisks

      Letters_or_Group <- "star_group"

    }

    # Pull the columns corresponding to the given Graph_Stat_Test

    if (Graph_Stat_Test == "Tukey") {

      # Add the Stat_Test name to the Letters_or_group information

      Letters_or_Group <- paste("Tukey_", Letters_or_Group, sep = "")

      Statistics_for_Graph <- apply(expand.grid(
        Letters_or_Group, Structures_list),
        MARGIN = 1,
        paste0,
        collapse = "_")

    } else if (Graph_Stat_Test == "Dunn") {

      # Add the Stat_Test name to the Letters_or_group information

      Letters_or_Group <- paste("Dunn_", Letters_or_Group, sep = "")

      Statistics_for_Graph <- apply(expand.grid(
        Letters_or_Group, Structures_list),
        MARGIN = 1,
        paste0,
        collapse = "_")

    } else if (Graph_Stat_Test == "Wilcoxon") {

      # Add the Stat_Test name to the Letters_or_group information

      Letters_or_Group <- paste("Wilcoxon_", Letters_or_Group, sep = "")

      Statistics_for_Graph <- apply(expand.grid(
        Letters_or_Group, Structures_list),
        MARGIN = 1,
        paste0,
        collapse = "_")

    }

    # Extract the necessary information from Statistics_Results
    # Rename columns to conform to the Letter codes for each Structure
    # Pivot according to Structures
    # Add information about the Stat_position
    # Rename the default column name "value" as "Significance group"
    # Re-order the factors of Structures to a sensible order
    # Re-order factors of Condition according to the...
    # Non_Duplicated_Order_list as determined by the user
    # Replace the letter descriptions of the Structures with their full name

    Geom_text_information <- Statistics_Results %>%
      select(all_of(c("Condition",
                      "Facet_1",
                      "Facet_2",
                      "Facet_3",
                      Statistics_for_Graph))) %>%
      rename_with(~ gsub('.*group_', '', .x)) %>%
      pivot_longer(cols = all_of(Structures),
                   names_to = "variable") %>%
      mutate(y_position = Stat_position) %>%
      mutate(Significance_Group = value) %>%
      mutate(variable = factor(variable,
                               levels = Structures)) %>%
      mutate(Condition = factor(
        Condition,
        levels = Non_Duplicated_Condition_Order_list)) %>%
      mutate(variable = revalue(
        variable,
        c("Total" = "Total",
          "EH"    = "Extraradical hyphae",
          "H"     = "Hyphopodia",
          "IH"    = "Intraradical hyphae",
          "A"     = "Arbuscules",
          "V"     = "Vesicles",
          "S"     = "Spores"
        ), warn_missing = FALSE
      ))

  }

  # Apply the Facet_X_Order user input (if present). Both Plot_Dataset and...
  # Geom_text_information (if it exists) must be modified.

  if (!is.null(Graph_Facet_1_Order)) {

    Plot_Dataset <- Plot_Dataset %>%
      mutate(Facet_1 = factor(Facet_1, levels = Graph_Facet_1_Order))

    if (exists("Geom_text_information")) {
      Geom_text_information <- Geom_text_information %>%
        mutate(Facet_1 = factor(Facet_1, levels = Graph_Facet_1_Order))

    }

  }

  if (!is.null(Graph_Facet_2_Order)) {

    Plot_Dataset <- Plot_Dataset %>%
      mutate(Facet_2 = factor(Facet_2, levels = Graph_Facet_2_Order))

    if (exists("Geom_text_information")) {
      Geom_text_information <- Geom_text_information %>%
        mutate(Facet_2 = factor(Facet_2, levels = Graph_Facet_2_Order))

    }

  }

  if (!is.null(Graph_Facet_3_Order)) {

    Plot_Dataset <- Plot_Dataset %>%
      mutate(Facet_3 = factor(Facet_3, levels = Graph_Facet_3_Order))

    if (exists("Geom_text_information")) {
      Geom_text_information <- Geom_text_information %>%
        mutate(Facet_3 = factor(Facet_3, levels = Graph_Facet_3_Order))

    }

  }

  # Save data to global environment for user

  assign("Plot_Dataset", Plot_Dataset, envir=.GlobalEnv)

  if (exists("Geom_text_information")) {
    assign("Geom_text_information", Geom_text_information, envir=.GlobalEnv)
  }

  #####
  # I) Set required values for ggplot components
  #####

  # Some sizes are dependent on the number of Conditions (the length of the...
  # x-axis)

  Number_of_conditions <- length(unique(Filtered_Dataset$Condition))

  # The size of the "Percentage Colonisation (%)" Y-label

  Y_Axis_Size <- 8 * Graph_Size_Y_Axis

  # The size of the "Condition" X-label

  X_Axis_Size <- 5 * Graph_Size_X_Axis

  # The size of the text for the figure legend

  Legend_Text_Size <- 6 * Graph_Size_Legend_Text

  # The size of the percentages on the Y-axis

  Numbers_Size <- 6 * Graph_Size_Percentages

  # The size of the labels giving statistics information

  if (Graph_Stat_Display == "Letters") {
    Geom_Text_Size <- 1.8 * Graph_Size_Statistics
  } else if (Graph_Stat_Display == "Reference") {
    Geom_Text_Size <- 4 * Graph_Size_Statistics
  }

  # The size of the figure legend

  Graph_Legend_Size <- 10 * Graph_Size_Legend

  # The size of the "Total", "Extraradical hyphae" etc. labels when plotting...
  # multiple facets.

  Right_Label_Size <- 8 * Graph_Size_Right_Label

  # The size of the labels for each facet on the top of the plot when...
  # plotting multiple facets
  # This will depend on the number of Facets being displayed
  # Facet_X_Data, from earlier, found the unique variables for each Facet.
  # If the facet is not being plotted, this needs to be ignored, set to NULL.

  if (Facet_1 == FALSE) {
    Length_Facet_1 = 1
  } else if (Facet_1 == TRUE) {
    Length_Facet_1 <- length(Facet_1_Data)
  }

  if (Facet_2 == FALSE) {
    Length_Facet_2 = 1
  } else if (Facet_2 == TRUE) {
    Length_Facet_2 <- length(Facet_2_Data)
  }

  if (Facet_3 == FALSE) {
    Length_Facet_3 = 1
  } else if (Facet_3 == TRUE) {
    Length_Facet_3 <- length(Facet_3_Data)
  }

  Number_of_facets <- Length_Facet_1 * Length_Facet_2 * Length_Facet_3

  # If there are no facets used at all, (Facet_1 = FALSE, Facet_2 = FALSE,
  # Facet_3 = FALSE) and Graph_Type = 'Facets' then the number of facets is
  # determined by the Structures being displayed

  if ((Facet_1 == FALSE)
      & (Facet_2 == FALSE)
      & (Facet_3 == FALSE)
      & (Graph_Type == "Facets")) {
    Number_of_facets <- length(Structures)
  }

  if (Number_of_facets <= 2) {
    Top_Label_Size <- 12 * Graph_Size_Top_Label
  } else if (Number_of_facets <= 4) {
    Top_Label_Size <- 8 * Graph_Size_Top_Label
  } else if (Number_of_facets <= 8) {
    Top_Label_Size <- 6 * Graph_Size_Top_Label
  } else if (Number_of_facets <= 12) {
    Top_Label_Size <- 4 * Graph_Size_Top_Label
  } else {
    Top_Label_Size <- 3 * Graph_Size_Top_Label
  }

  # The size of the data points overlayed on the bar chart

  if (Graph_Type == "Facets") {
    Point_size <- 0.5 * Graph_Size_Datapoints
  } else if (Graph_Type == "Single") {
    Point_size <- 0.2 * Graph_Size_Datapoints
  }

  # The overall width of the graph space

  Graph_Width <- (Number_of_conditions) + 3

  # The overall height of the graph space

  if ((length(Facets_Include) == 0)
      | (Graph_Type == "Single")) {
    Graph_Height <- 3
  } else {
    Graph_Height <- length(Structures) * 2
  }

  #####
  # J) Create the plot
  #####

  # The ggplot theme information

  custom_theme <- {
    theme(plot.width = Graph_Width) +
      theme(plot.height = Graph_Height) +
      theme_gray() +
      theme(
        panel.grid       = element_blank(),
        strip.background = element_rect(colour = "white",
                                        fill   = "white"),
        strip.text.x     = element_text(size   = Top_Label_Size,
                                        colour = Graph_Text_Colour,
                                        face   = "bold"),
        strip.text.y     = element_text(size   = Right_Label_Size,
                                        colour = Graph_Text_Colour,
                                        face   = "bold"),
        axis.title.x     = element_blank(),
        axis.title.y     = element_text(size   = Y_Axis_Size,
                                        face   = "bold",
                                        colour = Graph_Text_Colour
        ),
        axis.text.y      = element_text(size   = Numbers_Size,
                                        face   = "bold",
                                        colour = Graph_Text_Colour
        ),
        axis.text.x      = element_text(angle  = 45,
                                        vjust  = 1,
                                        hjust  = 1,
                                        family = "sans",
                                        colour = Graph_Text_Colour,
                                        size = X_Axis_Size
        ),
        panel.spacing.y  = unit(0.4, "lines"),
        legend.text      = element_text(size = Legend_Text_Size,
                                        face = "italic",
                                        color = Graph_Text_Colour),
        legend.title     = element_blank(),
        legend.key.size  = unit(Graph_Legend_Size, "points"),
        panel.background = element_rect(fill = Graph_Background_Colour),

      )
  }

  # Build the common features of each plot

  Plot <- Plot_Dataset %>%
    ggplot(mapping = aes(x = Condition, y = value)) +
    custom_theme +
    scale_y_continuous(limits = Coordinates,
                       expand = c(0,0)) +
    geom_hline(yintercept = c(25, 50, 75, 100),
               linewidth = 0.5,
               alpha = 0.3,
               color = Graph_Hline_Colour) +
    ylab("Percentage Colonisation (%)")

  # Add data

  if (Graph_Type == "Facets") {

    Plot <- Plot +
      stat_summary(mapping = aes(fill = Condition),
                   fun = mean,
                   geom = "bar",
                   position = position_dodge(width = 0.9),
                   colour = NA,
                   linewidth = 0,
                   alpha = 1) +
      geom_point(shape = 19,
                 size = Point_size,
                 colour = "black")

  } else if (Graph_Type == "Single") {

    Plot <- Plot +
      stat_summary(mapping = aes(fill = variable, group = variable),
                   fun = mean,
                   geom = "bar",
                   position = position_dodge(width = 0.9),
                   colour = NA,
                   linewidth = 0,
                   alpha = 1) +
      geom_point(position=position_dodge(width=0.90),
                 aes(Condition, value, group=variable),
                 shape = 19,
                 size = Point_size,
                 colour = "black")
  }

  # Create the Facet_Formula as defined by Facets_Include from user inputs
  # Form of variable ~ Facet_1 + Facet_2 + Facet_3 for Graph_Type == "Facets"
  # Form of ~ Facet_1 + Facet_2 + Facet_3 for Graph_Type == "Single

  if (length(Facets_Include > 0)) {
    if (Graph_Type == "Facets") {
      Facet_Formula <- as.formula(paste("variable",
                                        paste(Facets_Include, collapse=" + "),
                                        sep=" ~ "))

    } else if (Graph_Type == "Single") {
      Facet_Formula <- as.formula(paste(" ",
                                        paste(Facets_Include, collapse=" + "),
                                        sep=" ~ "))
    }
  }

  # Apply the chosen faceting rule

  if ((Graph_Type == "Facets")
      & (length(Facets_Include) == 0)) {
    Plot <- Plot +
      facet_grid( ~ variable, labeller = label_wrap_gen(width = 10),
                  scales = "free_x", space = "free_x")

  } else if ((Graph_Type == "Facets")
             & (length(Facets_Include) > 0)) {
    Plot <- Plot +
      facet_grid(Facet_Formula, labeller = label_wrap_gen(width = 10),
                 scales = "free_x", space = "free_x")

  } else if ((Graph_Type == "Single")
             & (length(Facets_Include) == 0)) {

    # Do not add a facet

  } else if ((Graph_Type == "Single")
             & (length(Facets_Include) > 0)) {
    Plot <- Plot +
      facet_grid(Facet_Formula, labeller = label_wrap_gen(width = 10),
                 scales = "free_x", space = "free_x")
  }

  # If there are no facets, and the significance group labels have more than...
  # two characters, then to reduce the risk of geom_text overlapping...
  # give an angle of 45 degrees. This only applies if Graph_Stat_Display...
  # is "Letters"

  if (exists("Geom_text_information")) {
    Significance_Group_length <- max(
      nchar(Geom_text_information$Significance_Group, type = "chars"))
  } else {
    Significance_Group_length <- 0
  }

  if ((Graph_Stat_Display == "Letters")
      & (length(Facets_Include) == 0)
      & (Significance_Group_length > 2))  {
    Angle = 45
  } else {
    Angle = 0
  }

  # Apply the Graph_Stat_Test input, which is only applicable for...
  # Graph_Type == "Facets"

  if ((!is.null(Graph_Stat_Test))
      & (Graph_Type == "Facets")) {

    Plot <- Plot +
      geom_text(aes(label = Significance_Group,
                    y = y_position),
                data = Geom_text_information,
                size = Geom_Text_Size,
                angle = Angle)

  }

  # Add one of the available colour schemes.

  # Graph_Manual_Colour == TRUE overrides any other colour inputs
  # Otherwise, add the given colour scheme and palette

  if (Graph_Manual_Colour == TRUE) {
    col <- as.character(Plot_Dataset$Manual_Colour)
    names(col) <- as.character(Plot_Dataset$Condition)

    Plot <- Plot +
      scale_fill_manual(values = col)

  } else if (Graph_Colour == "Viridis") {
    Plot <- Plot +
      scale_fill_viridis_d(name="", option = Graph_Palette,
                           begin = Graph_Palette_Begin,
                           end = Graph_Palette_End)

  } else if (Graph_Colour == "Brewer") {
    Plot <- Plot +
      scale_fill_brewer(name="", palette = Graph_Palette)

  } else if (Graph_Colour == "Grey") {
    Plot <- Plot +
      scale_fill_grey(start = Graph_Palette_Begin, end = Graph_Palette_End)

  }

  # Allow the option to remove the figure legend

  if (Graph_Legend == "FALSE") {
    Plot <- Plot +
      theme(legend.position = "none")
  }

  # Save Plot to global environment for user

  assign("Plot", Plot, envir=.GlobalEnv)

  # Display Plot

  print(Plot)

  # Save the plot if Graph_Output is TRUE

  if (Graph_Output == TRUE) {
    ggsave(glue("{Path}/{Graph_File}"),
           height = Graph_Height_Adjustment * Graph_Height,
           width = Graph_Width_Adjustment * Graph_Width,
           limitsize = FALSE,
           dpi = Graph_Resolution)
  }

}
