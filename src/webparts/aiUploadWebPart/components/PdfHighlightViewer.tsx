import * as React from 'react';
import { IOcrPageResult } from '../services/IPdfOcr';
import { joinOcrWords, wordIndexAtPoint, wordsInRect, wordsInTextRange } from '../services/ocrSelection';
import styles from './AiUpload.module.scss';

export interface IPdfHighlightViewerProps {
  page: IOcrPageResult;
  selectedIndexes: number[];
  showStyles?: boolean;
  onSelectText: (text: string, indexes: number[]) => void;
}

interface IPdfHighlightViewerState {
  displayWidth: number;
  displayHeight: number;
  isDragging: boolean;
  startX: number;
  startY: number;
  currentX: number;
  currentY: number;
}

export default class PdfHighlightViewer extends React.Component<IPdfHighlightViewerProps, IPdfHighlightViewerState> {
  private _imageRef: React.RefObject<HTMLImageElement>;
  private _dragging: boolean;
  private _startX: number;
  private _startY: number;

  public constructor(props: IPdfHighlightViewerProps) {
    super(props);
    this._imageRef = React.createRef<HTMLImageElement>();
    this._dragging = false;
    this._startX = 0;
    this._startY = 0;
    this.state = {
      displayWidth: 0,
      displayHeight: 0,
      isDragging: false,
      startX: 0,
      startY: 0,
      currentX: 0,
      currentY: 0
    };
  }

  public componentDidMount(): void {
    window.addEventListener('resize', this._syncSize);
    window.addEventListener('mousemove', this._onWindowMouseMove);
    window.addEventListener('mouseup', this._onWindowMouseUp);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('resize', this._syncSize);
    window.removeEventListener('mousemove', this._onWindowMouseMove);
    window.removeEventListener('mouseup', this._onWindowMouseUp);
  }

  public componentDidUpdate(prevProps: IPdfHighlightViewerProps): void {
    if (prevProps.page.imageUrl !== this.props.page.imageUrl) {
      this._syncSize();
    }
  }

  public render(): React.ReactElement<IPdfHighlightViewerProps> {
    const { page, selectedIndexes, showStyles } = this.props;
    const { displayWidth, displayHeight, isDragging, startX, startY, currentX, currentY } = this.state;
    const words = page.words || [];
    const scale = page.width > 0 && displayWidth > 0 ? displayWidth / page.width : 1;
    const highlightIndexes = isDragging
      ? wordsInTextRange(words, startX, startY, currentX, currentY)
      : selectedIndexes;

    return (
      <div className={styles.pdfViewer}>
        <div
          className={styles.pageStage}
          onMouseDown={this._onMouseDown}
        >
          <img
            ref={this._imageRef}
            className={styles.pageImage}
            src={page.imageUrl}
            alt=""
            draggable={false}
            onLoad={this._syncSize}
          />
          {displayWidth > 0 && (
            <svg
              className={styles.highlightOverlay}
              width={displayWidth}
              height={displayHeight}
            >
              {words.map((word, index) => {
                const isSelected = highlightIndexes.indexOf(index) >= 0;
                const x = word.x0 * scale;
                const y = word.y0 * scale;
                const width = Math.max(1, (word.x1 - word.x0) * scale);
                const height = Math.max(1, (word.y1 - word.y0) * scale);
                return (
                  <g key={`${word.x0}-${word.y0}-${index}`}>
                    {showStyles && word.bold && !isSelected && (
                      <rect
                        x={x}
                        y={y}
                        width={width}
                        height={height}
                        className={styles.wordBold}
                      />
                    )}
                    <rect
                      x={x}
                      y={y}
                      width={width}
                      height={height}
                      className={isSelected ? styles.wordSelected : styles.wordHit}
                    />
                    {showStyles && word.underline && (
                      <line
                        x1={x}
                        x2={x + width}
                        y1={y + height}
                        y2={y + height}
                        className={styles.wordUnderline}
                      />
                    )}
                  </g>
                );
              })}
            </svg>
          )}
        </div>
      </div>
    );
  }

  private _syncSize = (): void => {
    const image = this._imageRef.current;
    if (!image) {
      return;
    }
    this.setState({
      displayWidth: image.clientWidth,
      displayHeight: image.clientHeight
    });
  };

  private _toImagePoint = (event: React.MouseEvent<HTMLDivElement> | MouseEvent): { x: number; y: number } | undefined => {
    const image = this._imageRef.current;
    const { page } = this.props;
    if (!image || page.width === 0 || image.clientWidth === 0) {
      return undefined;
    }
    const bounds = image.getBoundingClientRect();
    const scaleX = page.width / bounds.width;
    const scaleY = page.height / bounds.height;
    const x = (event.clientX - bounds.left) * scaleX;
    const y = (event.clientY - bounds.top) * scaleY;
    return {
      x: Math.max(0, Math.min(page.width, x)),
      y: Math.max(0, Math.min(page.height, y))
    };
  };

  private _onMouseDown = (event: React.MouseEvent<HTMLDivElement>): void => {
    if (event.button !== 0) {
      return;
    }
    const point = this._toImagePoint(event);
    if (!point) {
      return;
    }
    event.preventDefault();
    this._dragging = true;
    this._startX = point.x;
    this._startY = point.y;
    this.setState({
      isDragging: true,
      startX: point.x,
      startY: point.y,
      currentX: point.x,
      currentY: point.y
    });
  };

  private _onWindowMouseMove = (event: MouseEvent): void => {
    if (!this._dragging) {
      return;
    }
    const point = this._toImagePoint(event);
    if (!point) {
      return;
    }
    this.setState({
      currentX: point.x,
      currentY: point.y
    });
  };

  private _onWindowMouseUp = (event: MouseEvent): void => {
    if (!this._dragging) {
      return;
    }
    this._dragging = false;
    const point = this._toImagePoint(event) || {
      x: this.state.currentX,
      y: this.state.currentY
    };
    this._finishSelection(point.x, point.y);
  };

  private _finishSelection = (endX: number, endY: number): void => {
    const { page, onSelectText } = this.props;
    const startX = this._startX;
    const startY = this._startY;
    const words = page.words || [];
    const movement = Math.max(Math.abs(endX - startX), Math.abs(endY - startY));

    let highlightIndexes: number[];
    let valueIndexes: number[];
    if (movement < 4) {
      const wordIndex = wordIndexAtPoint(words, endX, endY);
      highlightIndexes = wordIndex >= 0 ? [wordIndex] : [];
      valueIndexes = highlightIndexes;
    } else {
      highlightIndexes = wordsInTextRange(words, startX, startY, endX, endY);
      valueIndexes = wordsInRect(words, { x0: startX, y0: startY, x1: endX, y1: endY });
      if (valueIndexes.length === 0) {
        valueIndexes = highlightIndexes;
      }
    }

    this.setState({
      isDragging: false,
      currentX: endX,
      currentY: endY
    });

    if (valueIndexes.length > 0) {
      const selectedWords = valueIndexes.map((index) => words[index]).filter((word) => !!word);
      onSelectText(joinOcrWords(selectedWords), highlightIndexes);
    }
  };
}
