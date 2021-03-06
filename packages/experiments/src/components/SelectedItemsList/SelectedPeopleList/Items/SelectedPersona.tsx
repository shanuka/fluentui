import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  getId,
  styled,
  classNamesFunction,
  IStyleFunctionOrObject,
  css,
  EventGroup,
  IDisposable,
} from 'office-ui-fabric-react/lib/Utilities';
import { Persona, PersonaSize, IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { ISelectedItemProps } from '../../SelectedItemsList.types';
import { getStyles } from './SelectedPersona.styles';
import { ISelectedPersonaStyles, ISelectedPersonaStyleProps } from './SelectedPersona.types';
import { ITheme, IProcessedStyleSet } from 'office-ui-fabric-react/lib/Styling';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { IDragDropOptions } from 'office-ui-fabric-react/lib/utilities/dragdrop/interfaces';

const getClassNames = classNamesFunction<ISelectedPersonaStyleProps, ISelectedPersonaStyles>();

type ISelectedPersonaProps<TPersona> = ISelectedItemProps<TPersona> & {
  isValid?: (item: TPersona) => boolean;
  canExpand?: (item: TPersona) => boolean;
  getExpandedItems?: (item: TPersona) => TPersona[];

  /**
   * Call to provide customized styling that will layer on top of the variant rules.
   */
  styles?: IStyleFunctionOrObject<ISelectedPersonaStyleProps, ISelectedPersonaStyles>;

  /**
   * Theme for the component.
   */
  theme?: ITheme;
};

const DEFAULT_DROPPING_CSS_CLASS = 'is-dropping';

/**
 * A selected persona with support for item removal and expansion.
 *
 * To use the removal / expansion, bind isValid / canExpand /  getExpandedItems
 * when passing the onRenderItem to your SelectedItemsList
 */
const SelectedPersonaInner = React.memo(
  <TPersona extends IPersonaProps = IPersonaProps>(props: ISelectedPersonaProps<TPersona>) => {
    const {
      item,
      onRemoveItem,
      onItemChange,
      removeButtonAriaLabel,
      index,
      selected,
      isValid,
      canExpand,
      getExpandedItems,
      styles,
      theme,
      dragDropHelper,
      dragDropEvents,
      eventsToRegister,
    } = props;
    const itemId = getId();
    const [root, setRoot] = React.useState<HTMLElement | undefined>();
    const [dragDropSubscription, setDragDropSubscription] = React.useState<IDisposable>();
    const [isDropping, setIsDropping] = React.useState(false);
    const [droppingClassNames, setDroppingClassNames] = React.useState('');
    const [droppingClassName, setDroppingClassName] = React.useState('');
    const [isDraggable, setIsDraggable] = React.useState<boolean | undefined>(false);

    const _onRootRef = React.useCallback((div: HTMLDivElement) => {
      if (div) {
        // Need to resolve the actual DOM node, not the component.
        // The element itself will be used for drag/drop and focusing.
        setRoot(ReactDOM.findDOMNode(div) as HTMLElement);
      } else {
        setRoot(undefined);
      }
    }, []);

    const _updateDroppingState = (newValue: boolean, event: DragEvent) => {
      if (!newValue) {
        if (dragDropEvents!.onDragLeave) {
          dragDropEvents!.onDragLeave(item, event);
        }
      } else if (dragDropEvents!.onDragEnter) {
        setDroppingClassNames(dragDropEvents!.onDragEnter(item, event));
      }

      if (isDropping !== newValue) {
        setIsDropping(newValue);
      }
    };

    const dragDropOptions: IDragDropOptions = {
      eventMap: eventsToRegister,
      selectionIndex: index,
      context: { data: item, index: index },
      canDrag: dragDropEvents?.canDrag,
      canDrop: dragDropEvents?.canDrop,
      onDragStart: dragDropEvents?.onDragStart,
      updateDropState: _updateDroppingState,
      onDrop: dragDropEvents?.onDrop,
      onDragEnd: dragDropEvents?.onDragEnd,
      onDragOver: dragDropEvents?.onDragOver,
    };

    const events = new EventGroup(this);

    React.useEffect(() => {
      setDragDropSubscription(dragDropHelper?.subscribe(root as HTMLElement, events, dragDropOptions));
      return () => {
        dragDropSubscription?.dispose();
        setDragDropSubscription(undefined);
      };
    }, [dragDropHelper, dragDropSubscription, root, events, dragDropOptions]);

    React.useEffect(() => {
      setIsDraggable(dragDropEvents ? !!(dragDropEvents.canDrag && dragDropEvents.canDrop) : undefined);
    }, [dragDropEvents]);

    React.useEffect(() => {
      setDroppingClassName(isDropping ? droppingClassNames || DEFAULT_DROPPING_CSS_CLASS : '');
    }, [isDropping, droppingClassNames]);

    const onExpandClicked = React.useCallback(
      ev => {
        ev.stopPropagation();
        ev.preventDefault();
        if (onItemChange && getExpandedItems) {
          onItemChange(getExpandedItems(item), index);
        }
      },
      [onItemChange, getExpandedItems, item, index],
    );

    const onRemoveClicked = React.useCallback(
      ev => {
        ev.stopPropagation();
        ev.preventDefault();
        onRemoveItem && onRemoveItem();
      },
      [onRemoveItem],
    );

    const classNames: IProcessedStyleSet<ISelectedPersonaStyles> = React.useMemo(
      () =>
        getClassNames(styles, {
          isSelected: selected || false,
          isValid: isValid ? isValid(item) : true,
          theme: theme!,
          droppingClassName,
        }),
      // TODO: evaluate whether to add deps on `item` and `styles`
      // eslint-disable-next-line react-hooks/exhaustive-deps
      [selected, isValid, theme],
    );

    const coinProps = {};

    return (
      <div
        ref={_onRootRef}
        {...(typeof isDraggable === 'boolean'
          ? {
              'data-is-draggable': isDraggable, // This data attribute is used by some host applications.
              draggable: isDraggable,
            }
          : {})}
        onContextMenu={props.onContextMenu}
        onClick={props.onClick}
        className={css('ms-PickerPersona-container', classNames.personaContainer)}
        data-is-focusable={true}
        data-is-sub-focuszone={true}
        data-selection-index={index}
        role={'listitem'}
        aria-labelledby={'selectedItemPersona-' + itemId}
      >
        <div hidden={!canExpand || !canExpand(item) || !getExpandedItems}>
          <IconButton
            onClick={onExpandClicked}
            iconProps={{ iconName: 'Add', style: { fontSize: '14px' } }}
            className={css('ms-PickerItem-removeButton', classNames.expandButton)}
            styles={classNames.subComponentStyles.actionButtonStyles()}
            ariaLabel={removeButtonAriaLabel}
          />
        </div>
        <div className={css(classNames.personaWrapper)}>
          <div
            className={css('ms-PickerItem-content', classNames.itemContentWrapper)}
            id={'selectedItemPersona-' + itemId}
          >
            <Persona
              {...item}
              size={PersonaSize.size32}
              styles={classNames.subComponentStyles.personaStyles}
              coinProps={coinProps}
            />
          </div>
          <IconButton
            onClick={onRemoveClicked}
            iconProps={{ iconName: 'Cancel', style: { fontSize: '14px' } }}
            className={css('ms-PickerItem-removeButton', classNames.removeButton)}
            styles={classNames.subComponentStyles.actionButtonStyles()}
            ariaLabel={removeButtonAriaLabel}
          />
        </div>
      </div>
    );
  },
);

// export casting back to typeof inner to preserve generics.
export const SelectedPersona = styled<ISelectedPersonaProps<any>, ISelectedPersonaStyleProps, ISelectedPersonaStyles>(
  SelectedPersonaInner,
  getStyles,
  undefined,
  {
    scope: 'SelectedPersona',
  },
) as typeof SelectedPersonaInner;
