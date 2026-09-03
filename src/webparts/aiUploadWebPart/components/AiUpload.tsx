import * as React from 'react';
import {
  ChoiceGroup,
  DatePicker,
  DayOfWeek,
  DefaultButton,
  defaultDatePickerStrings,
  Dropdown,
  IconButton,
  IChoiceGroupOption,
  IDropdownOption,
  Label,
  Link,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import styles from './AiUpload.module.scss';
import type { IAiUploadProps } from './IAiUploadProps';
import * as strings from 'AiUploadWebPartStrings';
import { locFormat } from '../loc/locFormat';
import { PdfOcrService } from '../services/PdfOcrService';
import { IOcrPageResult, IOcrProgress } from '../services/IPdfOcr';
import PdfHighlightViewer from './PdfHighlightViewer';
import { DEFAULT_FORM_FIELDS, isNameField, isOrganizationField, isReceiverField, isRefNoField, isRegistrationNumberField, isRequiredField, isSenderField, isSubjectField, missingRequiredFields } from '../constants/defaultFormFields';
import { nameFromPdfFile } from '../constants/incomingName';
import {
  canonicalLeadingBl,
  isLeadingBlField,
  LEADING_BL_OPTIONS
} from '../constants/blSiteMap';
import {
  isProjectNumberField,
  isValidProjectNumber,
  projectNumberFromYourRef,
  sanitizeProjectNumber
} from '../constants/projectNumber';
import {
  canonicalSubProjectNumber,
  isSubProjectNumberField,
  SUB_PROJECT_NONE,
  SUB_PROJECT_NUMBER_OPTIONS
} from '../constants/subProjectNumber';
import {
  formatIssueDate,
  isIssueDateField,
  ISSUE_DATE_DISPLAY_LABEL,
  parseIssueDate,
  sanitizeIssueDate
} from '../constants/issueDate';
import {
  canonicalYesNo,
  isYesNoChoiceField,
  YES_NO_OPTIONS,
  YES_VALUE
} from '../constants/yesNo';
import { extractFieldValues, extractOurRefNo, extractYourRefNo } from '../services/fieldExtractor';
import { extractFieldsWithAi, isAiExtractionConfigured } from '../services/AiFieldExtractor';
import { analyzeSignature, asPersonName, extractOrganizationAboveAddressee, extractReceiverAboveDearSir, extractSubjectBelowDearSir } from '../services/signatureSender';
import { SharePointUploadService } from '../services/SharePointUploadService';
import {
  buildUploadFileUrl,
  fileNameFromFields,
  resolveUploadDestination
} from '../services/uploadDestination';

interface IFormField {
  id: string;
  label: string;
  value: string;
}

interface IAiUploadState {
  file: File | undefined;
  pages: IOcrPageResult[];
  currentPage: number;
  selectedWordIndexes: number[];
  fields: IFormField[];
  activeFieldId: string | undefined;
  isProcessing: boolean;
  progress: IOcrProgress | undefined;
  error: string | undefined;
  info: string | undefined;
  success: string | undefined;
  successUrl: string | undefined;
  warning: string | undefined;
  isUploading: boolean;
  uploadStatus: string | undefined;
  showRequiredErrors: boolean;
}

export default class AiUpload extends React.Component<IAiUploadProps, IAiUploadState> {
  private _fileInput: React.RefObject<HTMLInputElement>;
  private _nextFieldId: number;
  private _originalPdfBytes: Uint8Array | undefined;
  private _calendarOpen: boolean;
  private _calendarObserver: MutationObserver | undefined;
  private _originalFocus: ((this: HTMLElement, options?: FocusOptions) => void) | undefined;
  private _focusPatchTimer: number | undefined;

  public constructor(props: IAiUploadProps) {
    super(props);
    this._fileInput = React.createRef<HTMLInputElement>();
    this._nextFieldId = 1;
    this._originalPdfBytes = undefined;
    this._calendarOpen = false;
    const fields = this._fieldsFromConfig(props.formFields);
    this.state = {
      file: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields,
      activeFieldId: this._defaultActiveFieldId(fields),
      isProcessing: false,
      progress: undefined,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      warning: undefined,
      isUploading: false,
      uploadStatus: undefined,
      showRequiredErrors: false
    };
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this._onDocumentMouseDownCapture, true);
    document.addEventListener('click', this._onDocumentClickCapture, true);
    document.addEventListener('submit', this._onDocumentSubmitCapture, true);
  }

  public componentDidUpdate(prevProps: IAiUploadProps): void {
    if (prevProps.formFields !== this.props.formFields) {
      const fields = this._fieldsFromConfig(this.props.formFields, this.state.fields);
      this.setState({
        fields,
        activeFieldId: this._defaultActiveFieldId(fields)
      });
    }
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this._onDocumentMouseDownCapture, true);
    document.removeEventListener('click', this._onDocumentClickCapture, true);
    document.removeEventListener('submit', this._onDocumentSubmitCapture, true);
    this._stopCalendarObserver();
    this._disablePreventScrollFocus();
    this._revokePageUrls(this.state.pages);
  }

  public render(): React.ReactElement<IAiUploadProps> {
    const { hasTeamsContext } = this.props;
    const {
      file,
      pages,
      currentPage,
      selectedWordIndexes,
      fields,
      activeFieldId,
      isProcessing,
      progress,
      error,
      info,
      success,
      successUrl,
      warning,
      isUploading,
      uploadStatus,
      showRequiredErrors
    } = this.state;
    const busy = isProcessing || isUploading;
    const converted = pages.length > 0 && !isProcessing;
    const requiredMissing = missingRequiredFields(fields);
    const markRequired = converted || showRequiredErrors;
    const percent = progress ? Math.max(0, Math.min(100, progress.percent)) / 100 : 0;
    const currentPreview = pages.filter((page) => page.pageNumber === currentPage)[0];
    const hasFieldValues = fields.some((field) => field.value.length > 0);
    const destination = resolveUploadDestination(fields, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate
    });
    const uploadFileName = file ? fileNameFromFields(file, fields) : '';
    const destinationUrl = buildUploadFileUrl(destination, uploadFileName);
    const destinationLabel = !destination.siteUrl
      ? strings.UploadDestinationPending
      : destinationUrl;

    return (
      <section className={`${styles.aiUpload} ${hasTeamsContext ? styles.teams : ''}`}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="xLarge" className={styles.title}>{strings.WebPartTitle}</Text>
            <Text variant="medium" className={styles.subtitle}>{strings.WebPartSubtitle}</Text>
          </Stack>

          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              onDismiss={this._clearError}
              dismissButtonAriaLabel={strings.Dismiss}
            >
              {error}
            </MessageBar>
          )}

          {!error && markRequired && requiredMissing.length > 0 && (
            <MessageBar messageBarType={MessageBarType.error}>
              {locFormat(strings.RequiredFieldsPrompt, 'Please fill the required fields: {0}', requiredMissing.join(', '))}
            </MessageBar>
          )}

          {info && (
            <MessageBar
              messageBarType={MessageBarType.info}
              onDismiss={this._clearInfo}
              dismissButtonAriaLabel={strings.Dismiss}
            >
              {info}
            </MessageBar>
          )}

          {success && (
            <MessageBar
              messageBarType={MessageBarType.success}
              onDismiss={this._clearSuccess}
              dismissButtonAriaLabel={strings.Dismiss}
            >
              {success}
              {successUrl && (
                <span>
                  {' '}
                  <Link href={successUrl} target="_blank">{strings.OpenUploadedFile}</Link>
                </span>
              )}
            </MessageBar>
          )}

          {warning && (
            <MessageBar
              messageBarType={MessageBarType.warning}
              onDismiss={this._clearWarning}
              dismissButtonAriaLabel={strings.Dismiss}
            >
              {warning}
            </MessageBar>
          )}

          <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
            <Stack tokens={{ childrenGap: 6 }}>
              <Label>{strings.SelectPdfLabel}</Label>
              <input
                ref={this._fileInput}
                type="file"
                accept="application/pdf,.pdf"
                className={styles.hiddenFileInput}
                onChange={this._onFileChange}
                disabled={busy}
              />
              <DefaultButton
                iconProps={{ iconName: 'PDF' }}
                text={file ? file.name : strings.ChooseFile}
                onClick={this._openFilePicker}
                disabled={busy}
              />
            </Stack>

            <PrimaryButton
              iconProps={{ iconName: 'TextDocument' }}
              text={strings.ConvertButton}
              onClick={this._onConvert}
              disabled={!file || busy}
            />
          </Stack>

          {isProcessing && progress && (
            <ProgressIndicator
              label={progress.status}
              description={
                progress.totalPages > 0
                  ? locFormat(
                    strings.PageProgress,
                    'Page {0} of {1}',
                    String(progress.page),
                    String(progress.totalPages)
                  )
                  : undefined
              }
              percentComplete={percent}
            />
          )}

          <div className={styles.results}>
            <div className={styles.pane}>
              <div className={styles.paneHeader}>
                <Text className={styles.paneTitle}>{strings.FormFieldsLabel}</Text>
              </div>
              <div className={styles.fieldsBody}>
                <Text variant="small" className={styles.hint}>{strings.HighlightHint}</Text>
                {fields.map((field) => (
                  <div
                    key={field.id}
                    className={`${styles.fieldCard} ${field.id === activeFieldId ? styles.fieldCardActive : ''} ${this._isMissingRequired(field, markRequired) ? styles.fieldCardMissing : ''}`}
                    onClick={() => this._setActiveField(field.id)}
                  >
                    <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
                      {isLeadingBlField(field.label) ? (
                        <Dropdown
                          label={field.label}
                          selectedKey={field.value || undefined}
                          options={this._leadingBlOptions(field.value)}
                          onChange={(_event, option) => this._onFieldValueChange(field.id, option ? String(option.key) : '')}
                          onFocus={() => this._setActiveField(field.id)}
                          placeholder={strings.LeadingBlPlaceholder}
                          required={isRequiredField(field.label)}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                        />
                      ) : isSubProjectNumberField(field.label) ? (
                        <Dropdown
                          label={field.label}
                          selectedKey={canonicalSubProjectNumber(field.value)}
                          options={this._subProjectNumberOptions()}
                          onChange={(_event, option) => this._onFieldValueChange(field.id, option ? String(option.key) : SUB_PROJECT_NONE)}
                          onFocus={() => this._setActiveField(field.id)}
                          required={isRequiredField(field.label)}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                        />
                      ) : isYesNoChoiceField(field.label) ? (
                        <ChoiceGroup
                          label={field.label}
                          selectedKey={canonicalYesNo(field.value) || YES_VALUE}
                          options={this._yesNoOptions()}
                          onChange={(_event, option) => this._onFieldValueChange(field.id, option ? String(option.key) : YES_VALUE)}
                          onFocus={() => this._setActiveField(field.id)}
                          required={isRequiredField(field.label)}
                          className={styles.yesNoGroup}
                          styles={{
                            flexContainer: {
                              display: 'flex',
                              flexDirection: 'row',
                              columnGap: '16px'
                            }
                          }}
                        />
                      ) : isIssueDateField(field.label) ? (
                        <Stack className={styles.fieldInput}>
                          <Label>{ISSUE_DATE_DISPLAY_LABEL}</Label>
                          <DatePicker
                            value={parseIssueDate(field.value)}
                            onSelectDate={(date) => this._onIssueDateSelect(field.id, date)}
                            formatDate={(date) => date ? formatIssueDate(date) : ''}
                            parseDateFromString={(text) => parseIssueDate(text) || null}
                            placeholder={strings.IssueDatePlaceholder || 'dd/MM/yyyy'}
                            allowTextInput={true}
                            disableAutoFocus={true}
                            firstDayOfWeek={DayOfWeek.Monday}
                            strings={defaultDatePickerStrings}
                            isRequired={isRequiredField(field.label)}
                            ariaLabel={ISSUE_DATE_DISPLAY_LABEL}
                            isMonthPickerVisible={false}
                            calendarProps={{
                              showGoToToday: false
                            }}
                            calloutProps={{
                              setInitialFocus: false,
                              preventDismissOnResize: true,
                              onMouseDown: (event) => this._patchCalendarButtons(event.currentTarget),
                              onClick: (event) => event.preventDefault(),
                              layerProps: {
                                onLayerDidMount: this._onCalendarLayerMount,
                                onLayerWillUnmount: this._onCalendarLayerUnmount
                              }
                            }}
                            textField={{
                              description: strings.IssueDateDescription || 'Date format: dd/MM/yyyy',
                              errorMessage: this._requiredError(field, markRequired),
                              onFocus: () => this._setActiveField(field.id),
                              onKeyDown: (event) => {
                                if (event.key === 'Enter') {
                                  event.preventDefault();
                                }
                              }
                            }}
                          />
                        </Stack>
                      ) : (
                        <TextField
                          label={field.label}
                          value={field.value}
                          onChange={(_event, newValue) => this._onFieldValueChange(field.id, newValue || '')}
                          onFocus={() => this._setActiveField(field.id)}
                          placeholder={
                            isNameField(field.label)
                              ? strings.NamePlaceholder
                              : isRegistrationNumberField(field.label)
                                ? strings.RegistrationNumberPlaceholder
                                : isProjectNumberField(field.label)
                                  ? strings.ProjectNumberPlaceholder
                                  : strings.FieldPlaceholder
                          }
                          description={
                            isNameField(field.label)
                              ? strings.NameDescription
                              : isRegistrationNumberField(field.label)
                                ? strings.RegistrationNumberDescription
                                : isProjectNumberField(field.label)
                                  ? strings.ProjectNumberDescription
                                  : undefined
                          }
                          maxLength={isProjectNumberField(field.label) ? 8 : undefined}
                          readOnly={isRegistrationNumberField(field.label)}
                          required={isRequiredField(field.label)}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                        />
                      )}
                      {!isRegistrationNumberField(field.label) && !isYesNoChoiceField(field.label) && (
                      <IconButton
                        iconProps={{ iconName: 'Clear' }}
                        title={strings.ClearField}
                        ariaLabel={strings.ClearField}
                        onClick={(event) => {
                          event.stopPropagation();
                          this._onFieldValueChange(field.id, '');
                        }}
                      />
                      )}
                    </Stack>
                  </div>
                ))}
              </div>
            </div>

            <div className={styles.pane}>
              <div className={styles.paneHeader}>
                <Text className={styles.paneTitle}>{strings.PdfPreviewLabel}</Text>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                  <IconButton
                    iconProps={{ iconName: 'ChevronLeft' }}
                    title={strings.PreviousPage}
                    ariaLabel={strings.PreviousPage}
                    disabled={pages.length === 0 || currentPage <= 1}
                    onClick={this._onPreviousPage}
                  />
                  <Text>
                    {locFormat(
                      strings.PageProgress,
                      'Page {0} of {1}',
                      String(pages.length > 0 ? currentPage : 0),
                      String(pages.length)
                    )}
                  </Text>
                  <IconButton
                    iconProps={{ iconName: 'ChevronRight' }}
                    title={strings.NextPage}
                    ariaLabel={strings.NextPage}
                    disabled={pages.length === 0 || currentPage >= pages.length}
                    onClick={this._onNextPage}
                  />
                </Stack>
              </div>
              {currentPreview ? (
                <PdfHighlightViewer
                  page={currentPreview}
                  selectedIndexes={selectedWordIndexes}
                  onSelectText={this._onPdfSelectText}
                />
              ) : (
                <div className={styles.placeholder}>
                  <Text>{strings.PdfPreviewPlaceholder}</Text>
                </div>
              )}
            </div>
          </div>

          <DefaultButton
            iconProps={{ iconName: 'Clear' }}
            text={strings.ClearButton}
            onClick={this._onClear}
            disabled={busy || (!file && pages.length === 0 && !hasFieldValues)}
          />

          <div className={styles.uploadBar}>
            <Text variant="small" className={styles.hint}>{strings.UploadHint}</Text>
            <Text variant="small" className={styles.destination}>
              {strings.UploadDestinationLabel}:{' '}
              {destinationUrl && destination.siteUrl ? (
                <Link href={destinationUrl} target="_blank">{destinationUrl}</Link>
              ) : (
                destinationLabel
              )}
            </Text>
            {isUploading && uploadStatus && (
              <ProgressIndicator label={uploadStatus} />
            )}
            {file && (
            <PrimaryButton
              iconProps={{ iconName: 'Upload' }}
              text={strings.UploadButton}
              onClick={this._onUpload}
              disabled={busy}
            />
            )}
          </div>
        </Stack>
      </section>
    );
  }

  private _requiredError = (field: IFormField, markRequired: boolean): string | undefined => {
    if (!this._isMissingRequired(field, markRequired)) {
      return undefined;
    }
    return strings.RequiredFieldError;
  };

  private _isMissingRequired = (field: IFormField, markRequired: boolean): boolean => {
    return markRequired && isRequiredField(field.label) && !(field.value || '').trim();
  };

  private _defaultActiveFieldId = (fields: IFormField[]): string | undefined => {
    const fillable = fields.filter((field) =>
      !isNameField(field.label) && !isRegistrationNumberField(field.label)
    )[0];
    if (fillable) {
      return fillable.id;
    }
    return fields.length > 0 ? fields[0].id : undefined;
  };

  private _defaultFieldValue = (label: string): string => {
    if (isSubProjectNumberField(label)) {
      return SUB_PROJECT_NONE;
    }
    if (isYesNoChoiceField(label)) {
      return YES_VALUE;
    }
    return '';
  };

  private _fieldsFromConfig = (config: string, existing?: IFormField[]): IFormField[] => {
    const labels = (config || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
    const names = labels.length > 0 ? labels : DEFAULT_FORM_FIELDS;
    const valuesByLabel = new Map<string, string>();
    (existing || []).forEach((field) => {
      valuesByLabel.set(field.label, field.value);
    });
    return this._syncRegistrationFromName(names.map((label) => ({
      id: `field-${this._nextFieldId++}`,
      label,
      value: valuesByLabel.get(label) || this._defaultFieldValue(label)
    })));
  };

  private _setActiveField = (fieldId: string): void => {
    this.setState({ activeFieldId: fieldId, error: undefined });
  };

  private _isCalendarElement = (node: EventTarget | null): boolean => {
    if (!node || !(node instanceof Element)) {
      return false;
    }
    return !!node.closest('.ms-DatePicker-callout, .ms-Calendar, .ms-DatePicker');
  };

  private _patchCalendarButtons = (root?: ParentNode | null): void => {
    const scope = root || document;
    const buttons = scope.querySelectorAll('.ms-DatePicker-callout button, .ms-Calendar button');
    for (let i = 0; i < buttons.length; i++) {
      (buttons[i] as HTMLButtonElement).type = 'button';
    }
  };

  private _startCalendarObserver = (): void => {
    this._stopCalendarObserver();
    if (typeof MutationObserver === 'undefined') {
      return;
    }
    this._calendarObserver = new MutationObserver(() => {
      this._patchCalendarButtons();
    });
    this._calendarObserver.observe(document.body, { childList: true, subtree: true });
    this._patchCalendarButtons();
  };

  private _stopCalendarObserver = (): void => {
    if (this._calendarObserver) {
      this._calendarObserver.disconnect();
      this._calendarObserver = undefined;
    }
  };

  private _enablePreventScrollFocus = (): void => {
    if (this._originalFocus) {
      return;
    }
    this._originalFocus = HTMLElement.prototype.focus;
    const original = this._originalFocus;
    HTMLElement.prototype.focus = function (this: HTMLElement, options?: FocusOptions): void {
      original.call(this, { ...(options || {}), preventScroll: true });
    };
  };

  private _disablePreventScrollFocus = (): void => {
    if (this._focusPatchTimer !== undefined) {
      window.clearTimeout(this._focusPatchTimer);
      this._focusPatchTimer = undefined;
    }
    if (this._originalFocus) {
      HTMLElement.prototype.focus = this._originalFocus;
      this._originalFocus = undefined;
    }
  };

  private _onCalendarLayerMount = (): void => {
    this._calendarOpen = true;
    this._enablePreventScrollFocus();
    this._startCalendarObserver();
  };

  private _onCalendarLayerUnmount = (): void => {
    this._calendarOpen = false;
    this._stopCalendarObserver();
    this._focusPatchTimer = window.setTimeout(() => {
      this._disablePreventScrollFocus();
    }, 300);
  };

  private _onDocumentMouseDownCapture = (event: MouseEvent): void => {
    if (!this._isCalendarElement(event.target)) {
      return;
    }
    const button = (event.target as HTMLElement).closest('button');
    if (button) {
      (button as HTMLButtonElement).type = 'button';
    }
    this._patchCalendarButtons((event.target as HTMLElement).closest('.ms-DatePicker-callout, .ms-Calendar'));
  };

  private _onDocumentClickCapture = (event: MouseEvent): void => {
    if (!this._isCalendarElement(event.target)) {
      return;
    }
    const button = (event.target as HTMLElement).closest('button');
    if (button) {
      (button as HTMLButtonElement).type = 'button';
    }
  };

  private _onDocumentSubmitCapture = (event: Event): void => {
    if (!this._calendarOpen && !this._isCalendarElement(document.activeElement) && !this._isCalendarElement(event.target)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
  };

  private _onIssueDateSelect = (fieldId: string, date: Date | null | undefined): void => {
    this._enablePreventScrollFocus();
    this._onFieldValueChange(fieldId, date ? formatIssueDate(date) : '');
  };

  private _onFieldValueChange = (fieldId: string, value: string): void => {
    this.setState((prev) => {
      const target = prev.fields.filter((field) => field.id === fieldId)[0];
      if (target && isRegistrationNumberField(target.label)) {
        return {
          fields: prev.fields,
          activeFieldId: fieldId,
          showRequiredErrors: prev.showRequiredErrors,
          error: prev.error
        };
      }
      const fields = prev.fields.map((field) => {
        if (field.id !== fieldId) {
          return field;
        }
        return {
          ...field,
          value: this._normalizeFieldValue(field.label, value)
        };
      });
      return {
        fields: this._syncRegistrationFromName(fields),
        activeFieldId: fieldId,
        showRequiredErrors: false,
        error: undefined
      };
    });
  };

  private _applyPdfFileName = (fields: IFormField[], fileName?: string): IFormField[] => {
    const pdfName = fileName ? nameFromPdfFile(fileName) : '';
    const withName = fields.map((field) => (
      isNameField(field.label) && pdfName ? { ...field, value: pdfName } : field
    ));
    return this._syncRegistrationFromName(withName);
  };

  private _syncRegistrationFromName = (fields: IFormField[]): IFormField[] => {
    const nameField = fields.filter((field) => isNameField(field.label))[0];
    const nameValue = nameField ? nameField.value : '';
    return fields.map((field) => (
      isRegistrationNumberField(field.label) ? { ...field, value: nameValue } : field
    ));
  };

  private _normalizeFieldValue = (label: string, value: string): string => {
    if (isLeadingBlField(label)) {
      return canonicalLeadingBl(value);
    }
    if (isProjectNumberField(label)) {
      return sanitizeProjectNumber(value);
    }
    if (isSubProjectNumberField(label)) {
      return canonicalSubProjectNumber(value);
    }
    if (isReceiverField(label)) {
      return this._stripHonorifics(this._stripParentheses(value));
    }
    if (isSenderField(label)) {
      return this._stripParentheses(value);
    }
    if (isIssueDateField(label)) {
      return sanitizeIssueDate(value);
    }
    if (isYesNoChoiceField(label)) {
      return canonicalYesNo(value);
    }
    return value;
  };

  private _stripHonorifics = (value: string): string => {
    return (value || '')
      .replace(/^(?:(?:mr|mrs|ms|miss|dr|ir|prof(?:essor)?|engr?|sir|madam|mdm|mx|messrs)\b\.?\s*)+/i, '')
      .replace(/\s*(?:先生|女士|小姐|太太)\s*$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  };

  private _stripParentheses = (value: string): string => {
    let text = (value || '').trim();
    let previous = '';
    while (text !== previous) {
      previous = text;
      text = text
        .replace(/^[(\uFF08]\s*([\s\S]*?)\s*[)\uFF09]$/, '$1')
        .replace(/[(\uFF08][^)\uFF09]*[)\uFF09]/g, '')
        .replace(/[()\uFF08\uFF09]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
    return text;
  };

  private _fillActiveField = (text: string): void => {
    const trimmed = text.trim();
    if (!trimmed) {
      return;
    }
    const { activeFieldId, fields } = this.state;
    let targetId = activeFieldId || (fields.length > 0 ? fields[0].id : undefined);
    if (!targetId) {
      this.setState({ error: strings.SelectFieldFirst });
      return;
    }
    const target = fields.filter((field) => field.id === targetId)[0];
    if (target && isRegistrationNumberField(target.label)) {
      const nameField = fields.filter((field) => isNameField(field.label))[0];
      targetId = nameField ? nameField.id : targetId;
    }
    const fillTarget = fields.filter((field) => field.id === targetId)[0];
    if (fillTarget && isNameField(fillTarget.label)) {
      return;
    }
    const value = fillTarget ? this._normalizeFieldValue(fillTarget.label, trimmed) : trimmed;
    this.setState((prev) => ({
      fields: this._syncRegistrationFromName(prev.fields.map((field) => field.id === targetId ? { ...field, value } : field)),
      activeFieldId: targetId,
      error: undefined
    }));
  };

  private _yesNoOptions = (): IChoiceGroupOption[] => {
    return YES_NO_OPTIONS.map((name) => ({
      key: name,
      text: name
    }));
  };

  private _subProjectNumberOptions = (): IDropdownOption[] => {
    return SUB_PROJECT_NUMBER_OPTIONS.map((name) => ({
      key: name,
      text: name
    }));
  };

  private _leadingBlOptions = (currentValue?: string): IDropdownOption[] => {
    const options: IDropdownOption[] = LEADING_BL_OPTIONS.map((name) => ({
      key: name,
      text: name
    }));
    const current = (currentValue || '').trim();
    if (current && LEADING_BL_OPTIONS.indexOf(current) < 0) {
      options.unshift({
        key: current,
        text: `${current} (not a mapped site)`
      });
    }
    return options;
  };

  private _onPdfSelectText = (text: string, indexes: number[]): void => {
    this.setState({ selectedWordIndexes: indexes });
    this._fillActiveField(text);
  };

  private _openFilePicker = (): void => {
    this._fileInput.current?.click();
  };

  private _onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const selected = event.target.files && event.target.files.length > 0
      ? event.target.files[0]
      : undefined;

    if (!selected) {
      return;
    }

    if (!selected.name.toLowerCase().endsWith('.pdf') && selected.type !== 'application/pdf') {
      this._originalPdfBytes = undefined;
      this.setState({
        error: strings.InvalidFileType,
        file: undefined
      });
      event.target.value = '';
      return;
    }

    this._keepOriginalPdf(selected).catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _keepOriginalPdf = async (selected: File): Promise<void> => {
    this._revokePageUrls(this.state.pages);
    try {
      this._originalPdfBytes = new Uint8Array(await selected.arrayBuffer());
    } catch {
      this._originalPdfBytes = undefined;
      this.setState({ error: strings.InvalidFileType, file: undefined });
      return;
    }

    this.setState({
      file: selected,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      warning: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields: this._applyPdfFileName(this.state.fields, selected.name),
      showRequiredErrors: false
    });
  };

  private _onPreviousPage = (): void => {
    this.setState((prev) => ({
      currentPage: Math.max(1, prev.currentPage - 1),
      selectedWordIndexes: []
    }));
  };

  private _onNextPage = (): void => {
    this.setState((prev) => ({
      currentPage: Math.min(prev.pages.length, prev.currentPage + 1),
      selectedWordIndexes: []
    }));
  };

  private _onConvert = (): void => {
    this._runOcr().catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _onUpload = (): void => {
    this._uploadToSharePoint().catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _uploadToSharePoint = async (): Promise<void> => {
    const { file, fields } = this.state;
    if (!file) {
      this.setState({ error: strings.UploadNeedFile });
      return;
    }

    const requiredMissing = missingRequiredFields(fields);
    if (requiredMissing.length > 0) {
      this.setState({
        showRequiredErrors: true,
        error: locFormat(strings.UploadMissingFields, 'Fill these required fields before uploading: {0}', requiredMissing.join(', '))
      });
      return;
    }

    const destination = resolveUploadDestination(fields, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate
    });

    if (destination.missingFields.length > 0) {
      this.setState({
        error: locFormat(strings.UploadMissingFields, 'Fill these required fields before uploading: {0}', destination.missingFields.join(', '))
      });
      return;
    }

    if (destination.unrecognizedLeadingBl) {
      this.setState({
        error: locFormat(strings.UploadUnknownLeadingBl, 'Leading BL "{0}" does not match a business line site. Choose a listed business line.', destination.unrecognizedLeadingBl)
      });
      return;
    }

    const projectNumberField = fields.filter((field) => isProjectNumberField(field.label))[0];
    if (projectNumberField && !isValidProjectNumber(projectNumberField.value)) {
      this.setState({ error: strings.UploadInvalidProjectNumber });
      return;
    }

    this.setState({
      isUploading: true,
      error: undefined,
      success: undefined,
      successUrl: undefined,
      warning: undefined,
      showRequiredErrors: false,
      uploadStatus: strings.UploadStarting
    });

    try {
      const service = new SharePointUploadService(this.props.spHttpClient);
      const result = await service.uploadPdf(
        destination,
        file,
        fileNameFromFields(file, fields),
        (status) => {
          this.setState({ uploadStatus: this._localizeUploadStatus(status) });
        },
        this.props.currentWebUrl,
        fields.map((field) => ({ label: field.label, value: field.value })),
        this._originalPdfBytes
      );
      this.setState({
        isUploading: false,
        uploadStatus: undefined,
        success: locFormat(strings.UploadSucceeded, 'Uploaded {0}.', result.fileName),
        successUrl: result.fileUrl,
        warning: result.metadataError || undefined
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : strings.UploadFailed;
      this.setState({
        isUploading: false,
        uploadStatus: undefined,
        error: message
      });
    }
  };

  private _runOcr = async (): Promise<void> => {
    const { file } = this.state;
    if (!file) {
      return;
    }

    this._revokePageUrls(this.state.pages);
    this.setState({
      isProcessing: true,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      warning: undefined,
      showRequiredErrors: false,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      progress: {
        page: 0,
        totalPages: 0,
        percent: 0,
        status: strings.LoadingEngine
      }
    });

    try {
      const result = await PdfOcrService.extractText(
        this._originalPdfBytes ? this._originalPdfBytes.slice() : file,
        'eng',
        (progress) => {
          this.setState({
            progress: {
              ...progress,
              status: this._localizeStatus(progress.status)
            }
          });
        },
        (page) => {
          this.setState((prev) => ({
            pages: prev.pages.concat([page]),
            currentPage: page.pageNumber
          }));
        }
      );

      this.setState({
        progress: {
          page: result.pages.length,
          totalPages: result.pages.length,
          percent: 100,
          status: strings.ExtractingFields
        }
      });

      let filled: { fields: IFormField[]; info: string | undefined };
      try {
        filled = await this._fillFields(result.pages);
      } catch {
        filled = { fields: this.state.fields, info: undefined };
      }
      this.setState({
        pages: result.pages,
        currentPage: 1,
        isProcessing: false,
        progress: undefined,
        fields: filled.fields,
        error: undefined,
        info: filled.info
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : '';
      this.setState({
        isProcessing: false,
        progress: undefined,
        error: message && message !== 'Error' ? message : strings.OcrFailed
      });
    }
  };

  private _fillFields = async (pages: IOcrPageResult[]): Promise<{
    fields: IFormField[];
    error: string | undefined;
    info: string | undefined;
  }> => {
    const labels = this.state.fields.map((field) => field.label);
    const firstPage = pages && pages.length > 0 ? pages[0] : undefined;
    const closingPage = pages && pages.length > 0 ? pages[pages.length - 1] : undefined;
    let keywordValues: { [label: string]: string } = {};
    try {
      keywordValues = extractFieldValues(pages || [], labels);
    } catch {
      keywordValues = {};
    }
    let signature = await analyzeSignature(closingPage).catch(() => ({
      region: undefined,
      senderName: '',
      textBelow: ''
    }));
    if (firstPage && closingPage && firstPage.pageNumber !== closingPage.pageNumber) {
      signature = {
        ...signature,
        region: undefined
      };
    }
    let receiverName = '';
    let subjectText = '';
    let refNo = '';
    let projectNumber = '';
    let organization = '';
    try {
      receiverName = extractReceiverAboveDearSir(firstPage);
    } catch {
      receiverName = '';
    }
    try {
      organization = extractOrganizationAboveAddressee(firstPage);
    } catch {
      organization = '';
    }
    try {
      subjectText = await extractSubjectBelowDearSir(firstPage);
    } catch {
      subjectText = '';
    }
    try {
      refNo = extractOurRefNo(pages || []);
    } catch {
      refNo = '';
    }
    try {
      projectNumber = projectNumberFromYourRef(extractYourRefNo(pages || []));
    } catch {
      projectNumber = '';
    }
    let aiValues: { [label: string]: string } = {};
    let info: string | undefined;
    const aiConfig = {
      endpoint: this.props.azureOpenAiEndpoint || '',
      apiKey: this.props.azureOpenAiApiKey || '',
      deployment: this.props.azureOpenAiDeployment || '',
      apiVersion: this.props.azureOpenAiApiVersion || '2024-08-01-preview'
    };

    if (isAiExtractionConfigured(aiConfig)) {
      try {
        const ocrText = (pages || []).map((page) => page.text || '').join('\n');
        aiValues = await extractFieldsWithAi(ocrText, labels, aiConfig, firstPage, signature, receiverName, subjectText, refNo, organization);
      } catch {
        aiValues = {};
      }
    } else {
      info = strings.AiNotConfiguredHint;
    }

    let fields = this.state.fields;
    try {
      fields = this._applyPdfFileName(this.state.fields.map((field) => {
        if (isNameField(field.label)) {
          return field;
        }
        const aiValue = aiValues[field.label];
        const keywordValue = keywordValues[field.label];
        let value = '';
        if (isSenderField(field.label)) {
          value = signature.senderName || asPersonName(aiValue || '');
        } else if (isReceiverField(field.label)) {
          const aiReceiver = (aiValue || '').trim();
          const fromAi = aiReceiver && !/^dear\b/i.test(aiReceiver)
            ? aiReceiver.split(/\r?\n/)[0].trim()
            : '';
          value = receiverName || fromAi;
        } else if (isSubjectField(field.label)) {
          const aiSubject = (aiValue || '').trim();
          value = subjectText || aiSubject;
        } else if (isRefNoField(field.label)) {
          value = refNo || (aiValue || '').trim();
        } else if (isProjectNumberField(field.label)) {
          value = projectNumber || sanitizeProjectNumber(aiValue || '') || keywordValue || '';
        } else if (isOrganizationField(field.label)) {
          const aiOrganization = (aiValue || '').trim().split(/\r?\n/)[0].trim();
          value = organization || aiOrganization || keywordValue || '';
        } else {
          value = (aiValue && aiValue.trim()) || keywordValue || '';
        }
        value = this._normalizeFieldValue(field.label, value);
        return {
          ...field,
          value
        };
      }), this.state.file ? this.state.file.name : undefined);
    } catch {
      fields = this.state.fields;
    }
    return {
      error: undefined,
      info,
      fields
    };
  };

  private _localizeStatus = (status: string): string => {
    if (status.indexOf('Loading OCR engine') === 0) {
      return strings.LoadingEngine;
    }
    if (status.indexOf('Rendering page') === 0) {
      return status.replace('Rendering page', strings.RenderingPage).replace(' of ', ` ${strings.Of} `);
    }
    if (status.indexOf('OCR page') === 0) {
      return status.replace('OCR page', strings.OcrPage).replace(' of ', ` ${strings.Of} `);
    }
    if (status === 'Completed') {
      return strings.Completed;
    }
    if (status === strings.ExtractingFields) {
      return strings.ExtractingFields;
    }
    return status;
  };

  private _localizeUploadStatus = (status: string): string => {
    if (status.indexOf('Checking Root URL Mapping') === 0) {
      return strings.UploadCheckingMapping;
    }
    if (status.indexOf('Reading destination') === 0) {
      return strings.UploadReadingSite;
    }
    if (status.indexOf('Finding Project Documents') === 0) {
      return strings.UploadFindingLibrary;
    }
    if (status.indexOf('Creating folder') === 0) {
      return strings.UploadCreatingFolder;
    }
    if (status.indexOf('Uploading PDF') === 0) {
      return strings.UploadSendingFile;
    }
    if (status.indexOf('Updating library fields') === 0) {
      return strings.UploadUpdatingFields;
    }
    return status;
  };

  private _onClear = (): void => {
    if (this._fileInput.current) {
      this._fileInput.current.value = '';
    }
    this._originalPdfBytes = undefined;
    this._revokePageUrls(this.state.pages);
    this.setState({
      file: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields: this.state.fields.map((field) => ({ ...field, value: this._defaultFieldValue(field.label) })),
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      warning: undefined,
      isUploading: false,
      uploadStatus: undefined,
      showRequiredErrors: false,
      progress: undefined
    });
  };

  private _revokePageUrls = (pages: IOcrPageResult[]): void => {
    (pages || []).forEach((page) => {
      const url = page && page.imageUrl ? page.imageUrl : '';
      if (url.indexOf('blob:') === 0) {
        URL.revokeObjectURL(url);
      }
    });
  };

  private _clearError = (): void => {
    this.setState({ error: undefined });
  };

  private _clearInfo = (): void => {
    this.setState({ info: undefined });
  };

  private _clearSuccess = (): void => {
    this.setState({ success: undefined, successUrl: undefined });
  };

  private _clearWarning = (): void => {
    this.setState({ warning: undefined });
  };
}
