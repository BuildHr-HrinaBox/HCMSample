import React, { useState, useMemo, useRef, useEffect, useId } from 'react';
import './CityCombobox.css';

const MAX_VISIBLE_OPTIONS = 200;
export const COMBOBOX_OTHER_LABEL = 'Other';

function filterByPrefix(options, query, showAllWhenEmpty) {
  const q = String(query || '').trim().toLowerCase();
  if (!q) return showAllWhenEmpty ? options : [];
  return options.filter((c) => c.toLowerCase().startsWith(q));
}

function isListedOption(value, options) {
  const v = String(value || '').trim();
  if (!v) return false;
  return options.some((c) => c.toLowerCase() === v.toLowerCase());
}

/**
 * Searchable combobox with type-to-filter and "Other" for custom values (city, state, etc.).
 */
export function SearchableCombobox({
  id,
  name,
  value = '',
  options = [],
  onChange,
  onBlur,
  disabled = false,
  required = false,
  placeholder = 'Type to search',
  otherPlaceholder = 'Enter value',
  otherLabel = COMBOBOX_OTHER_LABEL,
  maxLength = 120,
  showAllWhenEmpty = false,
  emptyHint = 'Type a letter to search',
  noMatchEntity = 'items',
}) {
  const [open, setOpen] = useState(false);
  const [inputText, setInputText] = useState(value);
  const [otherMode, setOtherMode] = useState(() => {
    const v = String(value || '').trim();
    return v.length > 0 && !isListedOption(v, options);
  });
  const wrapperRef = useRef(null);
  const otherInputRef = useRef(null);
  const otherModeRef = useRef(otherMode);
  const inputTextRef = useRef(inputText);
  const otherValueRef = useRef(value);
  const blurTimerRef = useRef(null);
  const listId = useId();

  otherModeRef.current = otherMode;
  inputTextRef.current = inputText;
  otherValueRef.current = value;

  useEffect(() => {
    const v = String(value || '').trim();
    if (otherModeRef.current) {
      // In "Other" mode, only exit when the saved value is an exact list match.
      if (v && isListedOption(v, options)) {
        setOtherMode(false);
      }
    } else if (v && isListedOption(v, options)) {
      setOtherMode(false);
    } else if (v && !isListedOption(v, options)) {
      setOtherMode(true);
    }
    if (!open) setInputText(value);
  }, [value, options, open]);

  useEffect(
    () => () => {
      if (blurTimerRef.current) clearTimeout(blurTimerRef.current);
    },
    []
  );

  const trimmedQuery = inputText.trim();
  const filteredOptions = useMemo(
    () => (open && !otherMode ? filterByPrefix(options, inputText, showAllWhenEmpty) : []),
    [options, inputText, open, otherMode, showAllWhenEmpty]
  );

  const showList = open && !disabled && !otherMode;
  const totalMatches = filteredOptions.length;
  const visibleOptions = filteredOptions.slice(0, MAX_VISIBLE_OPTIONS);
  const hasMatches = totalMatches > 0;
  const hasMore = totalMatches > MAX_VISIBLE_OPTIONS;

  const emitChange = (nextValue) => {
    const v = String(nextValue ?? '').slice(0, maxLength);
    setInputText(v);
    onChange?.({ target: { name, value: v } });
  };

  const selectOption = (option) => {
    otherModeRef.current = false;
    setOtherMode(false);
    emitChange(option);
    setOpen(false);
  };

  const selectOther = () => {
    if (blurTimerRef.current) {
      clearTimeout(blurTimerRef.current);
      blurTimerRef.current = null;
    }
    setOpen(false);
    otherModeRef.current = true;
    setOtherMode(true);
    const typed = inputTextRef.current.trim();
    const keepTyped =
      typed.length > 0 && !options.some((c) => c.toLowerCase() === typed.toLowerCase());
    emitChange(keepTyped ? typed : '');
    window.requestAnimationFrame(() => {
      otherInputRef.current?.focus();
    });
  };

  const returnToList = () => {
    otherModeRef.current = false;
    setOtherMode(false);
    emitChange('');
    setInputText('');
    setOpen(false);
  };

  const handleFocus = () => {
    if (disabled || otherMode) return;
    setInputText(value || '');
    setOpen(true);
  };

  const handleInputChange = (e) => {
    const v = e.target.value.slice(0, maxLength);
    setInputText(v);
    inputTextRef.current = v;
    setOpen(true);
    onChange?.({ target: { name, value: v } });
  };

  const handleOtherInputChange = (e) => {
    const v = e.target.value.slice(0, maxLength);
    otherValueRef.current = v;
    emitChange(v);
  };

  const commitOnBlur = () => {
    if (otherModeRef.current) return;
    const trimmed = inputTextRef.current.trim();
    const exact = options.find((c) => c.toLowerCase() === trimmed.toLowerCase());
    if (exact) {
      otherModeRef.current = false;
      setOtherMode(false);
      emitChange(exact);
    } else if (!trimmed) {
      emitChange('');
    } else {
      otherModeRef.current = true;
      setOtherMode(true);
      emitChange(trimmed);
    }
  };

  const handleBlur = (e) => {
    const related = e.relatedTarget;
    if (related && wrapperRef.current?.contains(related)) return;
    if (blurTimerRef.current) clearTimeout(blurTimerRef.current);
    blurTimerRef.current = window.setTimeout(() => {
      blurTimerRef.current = null;
      if (otherModeRef.current) return;
      if (wrapperRef.current?.contains(document.activeElement)) return;
      setOpen(false);
      commitOnBlur();
      onBlur?.(e);
    }, 150);
  };

  const handleOtherBlur = (e) => {
    const latest = String(e.target?.value ?? otherValueRef.current ?? '').slice(0, maxLength);
    otherValueRef.current = latest;
    emitChange(latest);
    window.requestAnimationFrame(() => {
      onBlur?.(e);
    });
  };

  useEffect(() => {
    const onDocMouseDown = (e) => {
      if (!wrapperRef.current?.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, []);

  if (otherMode) {
    return (
      <div className="city-combobox city-combobox--other" ref={wrapperRef}>
        <input
          ref={otherInputRef}
          type="text"
          id={id}
          name={name}
          className="city-combobox-input"
          value={value}
          onChange={handleOtherInputChange}
          onBlur={handleOtherBlur}
          disabled={disabled}
          required={required}
          placeholder={otherPlaceholder}
          maxLength={maxLength}
          autoComplete="off"
        />
        {!disabled && (
          <button type="button" className="city-combobox-back" onClick={returnToList}>
            Choose from list
          </button>
        )}
      </div>
    );
  }

  return (
    <div className="city-combobox" ref={wrapperRef}>
      <input
        type="text"
        role="combobox"
        aria-expanded={showList}
        aria-controls={showList ? listId : undefined}
        aria-autocomplete="list"
        id={id}
        name={name}
        className="city-combobox-input"
        value={open ? inputText : value}
        onChange={handleInputChange}
        onFocus={handleFocus}
        onBlur={handleBlur}
        disabled={disabled}
        required={required}
        placeholder={placeholder}
        maxLength={maxLength}
        autoComplete="off"
      />
      {showList && (
        <ul id={listId} className="city-combobox-list" role="listbox">
          {!showAllWhenEmpty && trimmedQuery.length === 0 && (
            <li className="city-combobox-hint" aria-hidden="true">
              {emptyHint}
            </li>
          )}
          {hasMatches &&
            visibleOptions.map((option) => (
              <li key={option}>
                <button
                  type="button"
                  role="option"
                  className="city-combobox-option"
                  onMouseDown={(ev) => ev.preventDefault()}
                  onClick={() => selectOption(option)}
                >
                  {option}
                </button>
              </li>
            ))}
          {trimmedQuery.length > 0 && !hasMatches && (
            <li className="city-combobox-hint" role="status">
              No {noMatchEntity} starting with &ldquo;{trimmedQuery}&rdquo;
            </li>
          )}
          {hasMore && (
            <li className="city-combobox-more" aria-hidden="true">
              Showing {MAX_VISIBLE_OPTIONS} of {totalMatches} — type more letters to narrow
            </li>
          )}
          <li className="city-combobox-other-row">
            <button
              type="button"
              role="option"
              className="city-combobox-option city-combobox-option--other"
              onMouseDown={(ev) => {
                ev.preventDefault();
                selectOther();
              }}
            >
              {otherLabel}
            </button>
          </li>
        </ul>
      )}
    </div>
  );
}

/** City field — type a letter to filter, then pick or choose Other */
export default function CityCombobox(props) {
  return (
    <SearchableCombobox
      placeholder="Type letter to filter cities"
      otherPlaceholder="Enter your city name"
      emptyHint="Type a letter to search cities"
      noMatchEntity="cities"
      showAllWhenEmpty={false}
      {...props}
    />
  );
}

/** State field — same behaviour as City */
export function StateCombobox(props) {
  return (
    <SearchableCombobox
      placeholder="Type letter to filter states"
      otherPlaceholder="Enter your state name"
      emptyHint="Type a letter to search states"
      noMatchEntity="states"
      showAllWhenEmpty={false}
      {...props}
    />
  );
}

export const CITY_OTHER_LABEL = COMBOBOX_OTHER_LABEL;
