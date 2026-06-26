import React from 'react';
import Leave from './Leave';
import { NEW_LEAVE_API_BASE } from '../utils/leaveApi';

const NewLeave = (props) => (
  <Leave {...props} apiBase={NEW_LEAVE_API_BASE} pageTitle="New Leave" />
);

export default NewLeave;
