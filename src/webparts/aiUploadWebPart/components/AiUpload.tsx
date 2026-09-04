import * as React from 'react';
import {
  DatePicker,
  DayOfWeek,
  DefaultButton,
  defaultDatePickerStrings,
  Dropdown,
  IconButton,
  IDropdownOption,
  Label,
  Link,
  PrimaryButton,
  ProgressIndicator,
  TextField
} from '@fluentui/react';
import styles from './AiUpload.module.scss';
import type { IAiUploadProps } from './IAiUploadProps';
import * as strings from 'AiUploadWebPartStrings';
import { locFormat } from '../loc/locFormat';
import { PdfOcrService } from '../services/PdfOcrService';
import { IOcrPageResult, IOcrProgress } from '../services/IPdfOcr';
import { formatOcrTextWithStyles } from '../services/ocrSelection';
import PdfHighlightViewer from './PdfHighlightViewer';
import { DEFAULT_FORM_FIELDS, isNameField, isOrganizationField, isReceiverField, isRefNoField, isRegistrationNumberField, isRequiredField, isSenderField, isSubjectField, missingRequiredFields } from '../constants/defaultFormFields';
import { correspondenceKindFromFileName, nameFromPdfFile } from '../constants/incomingName';
import {
  canonicalLeadingBl,
  isLeadingBlField,
  LEADING_BL_OPTIONS
} from '../constants/blSiteMap';
import {
  isProjectNumberField,
  isValidProjectNumber,
  projectNumberFromRef,
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
  NO_VALUE,
  YES_NO_OPTIONS,
  YES_VALUE
} from '../constants/yesNo';
import { extractFieldValues, extractOurRefNo, extractOurRefOnly } from '../services/fieldExtractor';
import { extractFieldsWithAi, isAiExtractionConfigured } from '../services/AiFieldExtractor';
import { analyzeSignature, asPersonName, extractOrganizationAboveAddressee, extractReceiverAboveDearSir, extractSubjectBelowDearSir, subjectAppearsInPage } from '../services/signatureSender';
import { SharePointUploadService } from '../services/SharePointUploadService';
import {
  buildUploadFolderUrl,
  fileNameFromFields,
  resolveUploadDestination
} from '../services/uploadDestination';
import {
  rememberFieldValue,
  rememberFieldValues,
  isHistoryTextField,
  loadFieldHistory,
  saveFieldHistory,
  suggestionsFor,
  IFieldHistory
} from '../services/fieldHistory';
import { lookupLeadingBlFromNotificationSetup } from '../services/notificationSetup';

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
  successFolderUrl: string | undefined;
  warning: string | undefined;
  isUploading: boolean;
  uploadStatus: string | undefined;
  showRequiredErrors: boolean;
  showOcrStyles: boolean;
  history: IFieldHistory;
  historyFieldId: string | undefined;
}

export default class AiUpload extends React.Component<IAiUploadProps, IAiUploadState> {
  private _fileInput: React.RefObject<HTMLInputElement>;
  private _nextFieldId: number;
  private _originalPdfBytes: Uint8Array | undefined;
  private _calendarOpen: boolean;
  private _calendarObserver: MutationObserver | undefined;
  private _originalFocus: ((this: HTMLElement, options?: FocusOptions) => void) | undefined;
  private _focusPatchTimer: number | undefined;
  private _historyCloseTimer: number | undefined;
  private _leadingBlLookupSeq: number = 0;
  private _leadingBlLookupTimer: number | undefined;

  public constructor(props: IAiUploadProps) {
    super(props);
    this._fileInput = React.createRef<HTMLInputElement>();
    this._nextFieldId = 1;
    this._originalPdfBytes = undefined;
    this._calendarOpen = false;
    this._leadingBlLookupSeq = 0;
    this._leadingBlLookupTimer = undefined;
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
      successFolderUrl: undefined,
      warning: undefined,
      isUploading: false,
      uploadStatus: undefined,
      showRequiredErrors: false,
      showOcrStyles: false,
      history: loadFieldHistory(),
      historyFieldId: undefined
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
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
    }
    if (this._leadingBlLookupTimer !== undefined) {
      window.clearTimeout(this._leadingBlLookupTimer);
    }
    this._leadingBlLookupSeq = this._leadingBlLookupSeq + 1;
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
      successFolderUrl,
      warning,
      isUploading,
      uploadStatus,
      showRequiredErrors,
      showOcrStyles,
      historyFieldId
    } = this.state;
    const busy = isProcessing || isUploading;
    const converted = pages.length > 0 && !isProcessing;
    const requiredMissing = missingRequiredFields(fields);
    const markRequired = converted || showRequiredErrors;
    const percent = progress ? Math.max(0, Math.min(100, progress.percent)) / 100 : 0;
    const currentPreview = pages.filter((page) => page.pageNumber === currentPage)[0];
    const ocrInspectText = currentPreview
      ? (formatOcrTextWithStyles(currentPreview.words || []) || currentPreview.text || '')
      : '';
    const hasFieldValues = fields.some((field) => field.value.length > 0);
    const destination = resolveUploadDestination(fields, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate
    });
    const destinationUrl = buildUploadFolderUrl(destination);
    const destinationLabel = !destination.siteUrl
      ? strings.UploadDestinationPending
      : destinationUrl;
    const documentKind = file ? correspondenceKindFromFileName(file.name) : 'unknown';

    return (
      <section className={`${styles.aiUpload} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.header}>
          <h1 className={styles.title}>{strings.WebPartTitle}</h1>
          <p className={styles.subtitle}>{strings.WebPartSubtitle}</p>
        </div>

        {error && this._renderBanner('error', error, this._clearError)}
        {!error && markRequired && requiredMissing.length > 0 && this._renderBanner(
          'error',
          locFormat(strings.RequiredFieldsPrompt, 'Please fill the required fields: {0}', requiredMissing.join(', '))
        )}
        {info && this._renderBanner('info', info, this._clearInfo)}
        {success && (
          <div className={`${styles.banner} ${styles.bannerSuccess}`} role="status">
            <div className={styles.bannerBody}>
              <div className={styles.bannerText}>{success}</div>
              {(successUrl || successFolderUrl) && (
                <div className={styles.bannerActions}>
                  {successUrl && (
                    <a
                      className={styles.appleBtn}
                      href={successUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      {strings.OpenUploadedFile || 'Open file'}
                    </a>
                  )}
                  {successFolderUrl && (
                    <a
                      className={`${styles.appleBtn} ${styles.appleBtnSecondary}`}
                      href={successFolderUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      {strings.OpenUploadedFolder || 'Open folder'}
                    </a>
                  )}
                </div>
              )}
            </div>
            <button
              type="button"
              className={styles.bannerDismiss}
              onClick={this._clearSuccess}
              aria-label={strings.Dismiss}
            >
              ×
            </button>
          </div>
        )}
        {warning && this._renderBanner('warning', warning, this._clearWarning)}

        <div className={styles.toolbar}>
          <input
            ref={this._fileInput}
            type="file"
            accept="application/pdf,.pdf"
            className={styles.hiddenFileInput}
            onChange={this._onFileChange}
            disabled={busy}
          />
          <div className={styles.fileMeta}>
            <span className={styles.fileLabel}>{strings.SelectPdfLabel}</span>
            <span className={styles.fileNameRow}>
              <span className={styles.fileName}>{file ? file.name : strings.ChooseFile}</span>
              {documentKind !== 'unknown' && (
                <span className={`${styles.kindBadge} ${documentKind === 'incoming' ? styles.kindIncoming : styles.kindOutgoing}`}>
                  {documentKind === 'incoming'
                    ? (strings.IncomingLabel || 'Incoming')
                    : (strings.OutgoingLabel || 'Outgoing')}
                </span>
              )}
            </span>
          </div>
          <div className={styles.toolbarActions}>
            <DefaultButton
              className={styles.secondaryBtn}
              text={strings.ChooseFile}
              onClick={this._openFilePicker}
              disabled={busy}
            />
            <PrimaryButton
              className={styles.primaryBtn}
              text={strings.ConvertButton}
              onClick={this._onConvert}
              disabled={!file || busy}
            />
          </div>
        </div>

        {isProcessing && progress && (
          <div className={styles.progressWrap}>
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
          </div>
        )}

          <div className={styles.results}>
          <div className={styles.pane}>
            <div className={styles.paneHeader}>
              <span className={styles.paneTitle}>{strings.FormFieldsLabel}</span>
            </div>
            <div className={styles.fieldsBody}>
              <p className={styles.hint}>{strings.HighlightHint}</p>
              <div className={styles.fieldGroup}>
                {fields.map((field) => (
                  <div
                    key={field.id}
                    className={`${styles.fieldCard} ${field.id === activeFieldId ? styles.fieldCardActive : ''} ${this._isMissingRequired(field, markRequired) ? styles.fieldCardMissing : ''} ${field.id === historyFieldId ? styles.fieldCardHistoryOpen : ''}`}
                    onClick={() => this._setActiveField(field.id)}
                  >
                    {isYesNoChoiceField(field.label) ? (
                      this._renderYesNo(field)
                    ) : (
                    <div className={styles.fieldRow}>
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
                      ) : isIssueDateField(field.label) ? (
                        <div className={styles.fieldInput}>
                          <Label>{ISSUE_DATE_DISPLAY_LABEL}</Label>
                          <DatePicker
                            className={styles.datePicker}
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
                              showGoToToday: false,
                              styles: {
                                root: {
                                  width: 280
                                }
                              }
                            }}
                            calloutProps={{
                              className: styles.datePickerCallout,
                              gapSpace: 8,
                              isBeakVisible: false,
                              setInitialFocus: false,
                              preventDismissOnResize: true,
                              onMouseDown: (event) => this._patchCalendarButtons(event.currentTarget),
                              onClick: (event) => event.preventDefault(),
                              styles: {
                                root: {
                                  borderRadius: 16,
                                  overflow: 'hidden'
                                },
                                calloutMain: {
                                  borderRadius: 16,
                                  overflow: 'hidden'
                                }
                              },
                              layerProps: {
                                onLayerDidMount: this._onCalendarLayerMount,
                                onLayerWillUnmount: this._onCalendarLayerUnmount
                              }
                            }}
                            textField={{
                              borderless: true,
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
                        </div>
                      ) : (
                        <div className={styles.historyField}>
                        <TextField
                          label={field.label}
                          value={field.value}
                          onChange={(_event, newValue) => this._onFieldValueChange(field.id, newValue || '')}
                          onFocus={() => {
                            this._setActiveField(field.id);
                            this._openHistory(field.id);
                          }}
                          onBlur={() => this._onHistoryFieldBlur(field)}
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
                          borderless={true}
                        />
                        {this._renderFieldHistory(field)}
                        </div>
                      )}
                      {!isRegistrationNumberField(field.label) && (
                      <IconButton
                        className={styles.clearFieldBtn}
                        iconProps={{ iconName: 'Cancel' }}
                        title={strings.ClearField}
                        ariaLabel={strings.ClearField}
                        onClick={(event) => {
                          event.stopPropagation();
                          this._onFieldValueChange(field.id, '');
                        }}
                      />
                      )}
                    </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>

            <div className={styles.previewColumn}>
            <div className={styles.pane}>
              <div className={styles.paneHeader}>
                <span className={styles.paneTitle}>{strings.PdfPreviewLabel}</span>
                <div className={styles.pageNav}>
                  <IconButton
                    className={styles.pageBtn}
                    iconProps={{ iconName: 'ChevronLeft' }}
                    title={strings.PreviousPage}
                    ariaLabel={strings.PreviousPage}
                    disabled={pages.length === 0 || currentPage <= 1}
                    onClick={this._onPreviousPage}
                  />
                  <span className={styles.pageLabel}>
                    {locFormat(
                      strings.PageProgress,
                      'Page {0} of {1}',
                      String(pages.length > 0 ? currentPage : 0),
                      String(pages.length)
                    )}
                  </span>
                  <IconButton
                    className={styles.pageBtn}
                    iconProps={{ iconName: 'ChevronRight' }}
                    title={strings.NextPage}
                    ariaLabel={strings.NextPage}
                    disabled={pages.length === 0 || currentPage >= pages.length}
                    onClick={this._onNextPage}
                  />
                </div>
              </div>
              {currentPreview ? (
                <PdfHighlightViewer
                  page={currentPreview}
                  selectedIndexes={selectedWordIndexes}
                  showStyles={showOcrStyles}
                  onSelectText={this._onPdfSelectText}
                />
              ) : (
                <div className={styles.placeholder}>
                  {strings.PdfPreviewPlaceholder}
                </div>
              )}
            </div>

            <div className={styles.ocrTextPane}>
              <div className={styles.paneHeader}>
                <span className={styles.paneTitle}>{strings.ExtractedTextLabel}</span>
                <button
                  type="button"
                  className={`${styles.debugBtn} ${showOcrStyles ? styles.debugBtnOn : ''}`}
                  onClick={this._onToggleOcrStyles}
                >
                  Debug
                </button>
              </div>
              <div className={styles.ocrTextBody}>
                <p className={styles.hint}>{strings.ExtractedTextDescription}</p>
                <TextField
                  multiline={true}
                  readOnly={true}
                  resizable={true}
                  rows={8}
                  value={ocrInspectText}
                  placeholder={strings.ExtractedTextPlaceholder}
                  className={styles.ocrTextField}
                  borderless={true}
                />
              </div>
            </div>
            </div>
          </div>

          <div className={styles.footer}>
            <div className={styles.uploadBar}>
              <p className={styles.hint}>{strings.UploadHint}</p>
              <div className={styles.destination}>
                {strings.UploadDestinationLabel}:{' '}
                {destinationUrl && destination.siteUrl ? (
                  <Link href={destinationUrl} target="_blank">{destinationUrl}</Link>
                ) : (
                  destinationLabel
                )}
              </div>
              {isUploading && uploadStatus && (
                <ProgressIndicator label={uploadStatus} />
              )}
            </div>
            <div className={styles.footerActions}>
              <DefaultButton
                className={styles.ghostBtn}
                text={strings.ClearButton}
                onClick={this._onClear}
                disabled={busy || (!file && pages.length === 0 && !hasFieldValues)}
              />
              {file && (
                <PrimaryButton
                  className={styles.primaryBtn}
                  text={strings.UploadButton}
                  onClick={this._onUpload}
                  disabled={busy}
                />
              )}
            </div>
          </div>
      </section>
    );
  }

  private _pulseYesNo = (target: HTMLElement): void => {
    const group = target.closest(`.${styles.segmented}`) as HTMLElement | null;
    if (!group || typeof group.animate !== 'function') {
      return;
    }
    group.animate(
      [
        { transform: 'scale(1)' },
        { transform: 'scale(0.96)', offset: 0.32 },
        { transform: 'scale(1)' }
      ],
      { duration: 280, easing: 'cubic-bezier(0.22, 1, 0.36, 1)' }
    );
  };

  private _persistHistory = (history: IFieldHistory): IFieldHistory => {
    saveFieldHistory(history);
    return history;
  };

  private _openHistory = (fieldId: string): void => {
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
      this._historyCloseTimer = undefined;
    }
    this.setState({ historyFieldId: fieldId });
  };

  private _closeHistory = (): void => {
    this.setState({ historyFieldId: undefined });
  };

  private _onHistoryFieldBlur = (field: IFormField): void => {
    if (isHistoryTextField(field.label)) {
      const history = this._persistHistory(rememberFieldValue(this.state.history, field.label, field.value));
      this.setState({ history });
    }
    this._historyCloseTimer = window.setTimeout(() => {
      this._closeHistory();
    }, 160);
  };

  private _applyHistoryValue = (field: IFormField, value: string): void => {
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
      this._historyCloseTimer = undefined;
    }
    this._onFieldValueChange(field.id, value);
    this.setState({ historyFieldId: undefined });
  };

  private _renderFieldHistory = (field: IFormField): React.ReactNode => {
    if (field.id !== this.state.historyFieldId || !isHistoryTextField(field.label)) {
      return undefined;
    }
    const suggestions = suggestionsFor(this.state.history, field.label, field.value);
    if (suggestions.length === 0) {
      return undefined;
    }
    return (
      <ul className={styles.historyList} role="listbox" aria-label={strings.RecentValuesLabel}>
        <li className={styles.historyCaption}>{strings.RecentValuesLabel}</li>
        {suggestions.map((value) => (
          <li key={value}>
            <button
              type="button"
              className={styles.historyItem}
              onMouseDown={(event) => {
                event.preventDefault();
                event.stopPropagation();
                this._applyHistoryValue(field, value);
              }}
            >
              {value}
            </button>
          </li>
        ))}
      </ul>
    );
  };

  private _renderBanner = (
    kind: 'error' | 'info' | 'success' | 'warning',
    message: React.ReactNode,
    onDismiss?: () => void
  ): React.ReactNode => {
    const kindClass = kind === 'error'
      ? styles.bannerError
      : kind === 'info'
        ? styles.bannerInfo
        : kind === 'success'
          ? styles.bannerSuccess
          : styles.bannerWarning;
    return (
      <div className={`${styles.banner} ${kindClass}`} role="status">
        <div className={styles.bannerText}>{message}</div>
        {onDismiss && (
          <button
            type="button"
            className={styles.bannerDismiss}
            onClick={onDismiss}
            aria-label={strings.Dismiss}
          >
            ×
          </button>
        )}
      </div>
    );
  };

  private _renderYesNo = (field: IFormField): React.ReactNode => {
    const selected = canonicalYesNo(field.value) || YES_VALUE;
    const isNo = selected === NO_VALUE;
    return (
      <div className={styles.yesNoRow}>
        <span className={styles.yesNoLabel}>
          {field.label}
          {isRequiredField(field.label) ? <span className={styles.required}> *</span> : undefined}
        </span>
        <div
          className={`${styles.segmented} ${isNo ? styles.segmentedNo : ''}`}
          role="radiogroup"
          aria-label={field.label}
        >
          <span className={styles.segmentThumb} aria-hidden={true} />
          {YES_NO_OPTIONS.map((name) => (
            <button
              key={name}
              type="button"
              className={`${styles.segment} ${selected === name ? styles.segmentActive : ''}`}
              aria-checked={selected === name}
              role="radio"
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                this._pulseYesNo(event.currentTarget);
                this._setActiveField(field.id);
                this._onFieldValueChange(field.id, selected === YES_VALUE ? NO_VALUE : YES_VALUE);
              }}
            >
              {name}
            </button>
          ))}
        </div>
      </div>
    );
  };

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
    if (this._focusPatchTimer !== undefined) {
      window.clearTimeout(this._focusPatchTimer);
      this._focusPatchTimer = undefined;
    }
    this._enablePreventScrollFocus();
    this._startCalendarObserver();
  };

  private _onCalendarLayerUnmount = (): void => {
    this._calendarOpen = false;
    this._stopCalendarObserver();
    this._focusPatchTimer = window.setTimeout(() => {
      this._focusPatchTimer = undefined;
      if (!this._calendarOpen) {
        this._disablePreventScrollFocus();
      }
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
    const target = this.state.fields.filter((field) => field.id === fieldId)[0];
    this.setState((prev) => {
      const current = prev.fields.filter((field) => field.id === fieldId)[0];
      if (current && isRegistrationNumberField(current.label)) {
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
    if (target && isProjectNumberField(target.label)) {
      this._refreshLeadingBlFromNotificationSetup(this._normalizeFieldValue(target.label, value));
    }
  };

  private _refreshLeadingBlFromNotificationSetup = (projectNumber: string): void => {
    const file = this.state.file;
    if (!file || correspondenceKindFromFileName(file.name) === 'incoming') {
      return;
    }
    const projectNo = sanitizeProjectNumber(projectNumber);
    if (!isValidProjectNumber(projectNo)) {
      return;
    }
    if (this._leadingBlLookupTimer) {
      window.clearTimeout(this._leadingBlLookupTimer);
    }
    this._leadingBlLookupSeq = this._leadingBlLookupSeq + 1;
    const seq = this._leadingBlLookupSeq;
    this._leadingBlLookupTimer = window.setTimeout(() => {
      this._leadingBlLookupTimer = undefined;
      lookupLeadingBlFromNotificationSetup(this.props.spHttpClient, this.props.currentWebUrl, projectNo)
        .then((result) => {
          if (seq !== this._leadingBlLookupSeq) {
            return;
          }
          const warning = result.thresholdExceeded
            ? (strings.NotificationSetupThresholdHint ||
              'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).')
            : undefined;
          if (!result.leadingBl && !warning) {
            return;
          }
          this.setState((prev) => ({
            fields: result.leadingBl
              ? prev.fields.map((field) => (
                isLeadingBlField(field.label)
                  ? { ...field, value: canonicalLeadingBl(result.leadingBl) }
                  : field
              ))
              : prev.fields,
            warning: warning || prev.warning
          }));
        })
        .catch(() => {
          return;
        });
    }, 400);
  };

  private _applyPdfFileName = (fields: IFormField[], fileName?: string): IFormField[] => {
    const pdfName = fileName ? nameFromPdfFile(fileName) : '';
    const withName = fields.map((field) => (
      isNameField(field.label) && pdfName ? { ...field, value: pdfName } : field
    ));
    return this._syncRegistrationFromName(withName);
  };

  private _fieldsForSelectedFile = (fields: IFormField[], fileName: string): IFormField[] => {
    const kind = correspondenceKindFromFileName(fileName);
    const next = kind === 'incoming'
      ? fields.map((field) => (
        isNameField(field.label) || isRegistrationNumberField(field.label)
          ? field
          : { ...field, value: this._defaultFieldValue(field.label) }
      ))
      : fields;
    return this._applyPdfFileName(next, fileName);
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
      successFolderUrl: undefined,
      warning: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields: this._fieldsForSelectedFile(this.state.fields, selected.name),
      showRequiredErrors: false
    });
  };

  private _onToggleOcrStyles = (): void => {
    this.setState((prev) => ({
      showOcrStyles: !prev.showOcrStyles
    }));
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
      successFolderUrl: undefined,
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
        successFolderUrl: result.folderUrl,
        warning: result.metadataError || undefined,
        history: this._persistHistory(rememberFieldValues(this.state.history, fields))
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
      successFolderUrl: undefined,
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

      let filled: { fields: IFormField[]; info: string | undefined; warning?: string };
      const documentKind = correspondenceKindFromFileName(file.name);
      if (documentKind === 'incoming') {
        filled = {
          fields: this._applyPdfFileName(this.state.fields, file.name),
          info: strings.IncomingNoAutoFillHint || 'Incoming files are not auto-filled. Enter fields manually or highlight the PDF.'
        };
      } else {
        this.setState({
          progress: {
            page: result.pages.length,
            totalPages: result.pages.length,
            percent: 100,
            status: strings.ExtractingFields
          }
        });
        try {
          filled = await this._fillFields(result.pages);
        } catch {
          filled = { fields: this.state.fields, info: undefined };
        }
      }
      this.setState({
        pages: result.pages,
        currentPage: 1,
        isProcessing: false,
        progress: undefined,
        fields: filled.fields,
        error: undefined,
        info: filled.info,
        warning: filled.warning
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
    warning?: string;
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
      projectNumber = projectNumberFromRef(extractOurRefOnly(pages || []));
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
        const ocrText = (pages || []).map((page) => formatOcrTextWithStyles(page.words || []) || page.text || '').join('\n');
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
          const groundedAi = firstPage && aiSubject && subjectAppearsInPage(firstPage, aiSubject)
            ? aiSubject
            : '';
          value = subjectText || groundedAi;
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

    let warning: string | undefined;
    const projectField = fields.filter((field) => isProjectNumberField(field.label))[0];
    const resolvedProjectNumber = projectField ? sanitizeProjectNumber(projectField.value) : '';
    if (resolvedProjectNumber) {
      try {
        const setup = await lookupLeadingBlFromNotificationSetup(
          this.props.spHttpClient,
          this.props.currentWebUrl,
          resolvedProjectNumber
        );
        if (setup.leadingBl) {
          fields = fields.map((field) => (
            isLeadingBlField(field.label)
              ? { ...field, value: this._normalizeFieldValue(field.label, setup.leadingBl) }
              : field
          ));
        }
        if (setup.thresholdExceeded) {
          warning = strings.NotificationSetupThresholdHint ||
            'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).';
        }
      } catch {
        warning = undefined;
      }
    }

    return {
      error: undefined,
      info,
      warning,
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
      successFolderUrl: undefined,
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
    this.setState({ success: undefined, successUrl: undefined, successFolderUrl: undefined });
  };

  private _clearWarning = (): void => {
    this.setState({ warning: undefined });
  };
}
