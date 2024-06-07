import { databasePrefix } from './constants';
import { IInputs } from './generated/ManifestTypes';

export interface IBryntumSchedulerComponentProps {
  context?: ComponentFramework.Context<IInputs>;
}

type RecordItem = {
  data:
    | SchedulerEvent
    | SchedulerResource
    | SchedulerDependency
    | SchedulerAssignment;
  meta: {
    modified:
      | Partial<SchedulerEvent>
      | Partial<SchedulerResource>
      | Partial<SchedulerDependency>
      | Partial<SchedulerAssignment>;
  };
} & SchedulerEvent & SchedulerResource & SchedulerDependency & SchedulerAssignment;

export type SyncData = {
  action: 'dataset' | 'add' | 'remove' | 'update';
  records: RecordItem[];
  store: {
    id: 'events' | 'resources' | 'dependencies' | 'assignments';
  };
};

export type SchedulerEvent = {
  id: string;
  name: string;
  startDate: string;
  endDate: string;
  readOnly: number;
  timeZone: string;
  draggable: number;
  resizable: string | boolean;
  children: string;
  allDay: number;
  duration: number;
  durationUnit: string;
  exceptionDates: string;
  recurrenceRule: string;
  cls: string;
  eventColor: string;
  eventStyle: string;
  iconCls: string;
  style: string;
};

const KEY_ID = `${databasePrefix}id` as const;
const KEY_NAME = `${databasePrefix}name` as const;
const KEY_START_DATE = `${databasePrefix}startdate` as const;
const KEY_END_DATE = `${databasePrefix}enddate` as const;
const KEY_READONLY = `${databasePrefix}readonly` as const;
const KEY_TIME_ZONE = `${databasePrefix}timezone` as const;
const KEY_DRAGGABLE = `${databasePrefix}draggable` as const;
const KEY_RESIZABLE = `${databasePrefix}resizable` as const;
const KEY_CHILDREN = `${databasePrefix}children` as const;
const KEY_ALL_DAY = `${databasePrefix}allday` as const;
const KEY_DURATION = `${databasePrefix}duration` as const;
const KEY_DURATION_UNIT = `${databasePrefix}durationunit` as const;
const KEY_EXCEPTION_DATES = `${databasePrefix}exceptiondates` as const;
const KEY_RECURRENCE_RULE = `${databasePrefix}recurrencerule` as const;
const KEY_CLS = `${databasePrefix}cls` as const;
const KEY_EVENT_COLOR = `${databasePrefix}eventcolor` as const;
const KEY_EVENT_STYLE = `${databasePrefix}eventstyle` as const;
const KEY_ICON_CLS = `${databasePrefix}iconcls` as const;
const KEY_STYLE = `${databasePrefix}style` as const;

export type SchedulerEventDataverse = {
  [KEY_ID]: string;
  [KEY_NAME]: string;
  [KEY_START_DATE]: string;
  [KEY_END_DATE]: string;
  [KEY_READONLY]: boolean;
  [KEY_TIME_ZONE]: string;
  [KEY_DRAGGABLE]: boolean;
  [KEY_RESIZABLE]: string;
  [KEY_CHILDREN]: string;
  [KEY_ALL_DAY]: boolean;
  [KEY_DURATION]: number;
  [KEY_DURATION_UNIT]: string;
  [KEY_EXCEPTION_DATES]: string;
  [KEY_RECURRENCE_RULE]: string;
  [KEY_CLS]: string;
  [KEY_EVENT_COLOR]: string;
  [KEY_EVENT_STYLE]: string;
  [KEY_ICON_CLS]: string;
  [KEY_STYLE]: string;
};

export type SchedulerResource = {
  id: string;
  name: string;
  eventColor: string;
  readOnly: string;
};

export type SchedulerResourceDataverse = {
  [KEY_ID]: string;
  [KEY_NAME]: string;
  [KEY_EVENT_COLOR]: string;
  [KEY_READONLY]: boolean;
};

export type SchedulerDependency = {
  id: string;
  type: number;
  cls: string;
  lag: number;
  lagUnit: string;
  fromSide: string;
  toSide: string;
  from: string;
  to: string;
};

const KEY_TYPE = `${databasePrefix}type` as const;
const KEY_LAG = `${databasePrefix}lag` as const;
const KEY_LAG_UNIT = `${databasePrefix}lagunit` as const;
const KEY_FROM_SIDE = `${databasePrefix}fromside` as const;
const KEY_TO_SIDE = `${databasePrefix}toside` as const;
const KEY_FROM = `${databasePrefix}from@odata.bind` as const;
const KEY_TO = `${databasePrefix}to@odata.bind` as const;

export type SchedulerDependencyDataverse = {
  [KEY_ID]: string;
  [KEY_TYPE]: number;
  [KEY_CLS]: string;
  [KEY_LAG]: number;
  [KEY_LAG_UNIT]: string;
  [KEY_FROM_SIDE]: string;
  [KEY_TO_SIDE]: string;
  [KEY_FROM]: string;
  [KEY_TO]: string;
};

export type SchedulerAssignment = {
  id: string;
  eventId: string;
  resourceId: string;
};

const KEY_EVENT_ID = `${databasePrefix}eventid@odata.bind` as const;
const KEY_RESOURCE_ID = `${databasePrefix}resourceid@odata.bind` as const;

export type SchedulerAssignmentDataverse = {
  [KEY_ID]: string;
  [KEY_EVENT_ID]: string;
  [KEY_RESOURCE_ID]: string;
};