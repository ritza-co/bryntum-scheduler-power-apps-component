import { BryntumScheduler as OriginalBryntumScheduler } from "@bryntum/scheduler-react";
import * as React from "react";
import { FunctionComponent, useRef } from "react";
import {
  IBryntumSchedulerComponentProps,
  SchedulerAssignment,
  SchedulerAssignmentDataverse,
  SchedulerDependency,
  SchedulerDependencyDataverse,
  SchedulerEvent,
  SchedulerEventDataverse,
  SchedulerResource,
  SchedulerResourceDataverse,
  SyncData,
} from "./BryntumSchedulerComponent.types";
import LoadingSpinner from "./LoadingSpinner";
import {
  assignmentsDataverseTableName,
  databasePrefix,
  dependenciesDataverseTableName,
  eventsDataverseTableName,
  resourcesDataverseTableName,
} from "./constants";
import { schedulerConfig } from "./schedulerConfig";

function removeEventsDataColumnPrefixes(entities: any[]) {
  return entities.map((entity: ComponentFramework.WebApi.Entity) => {
    let event: Partial<SchedulerEvent> = {};
    if (entity[`${databasePrefix}${eventsDataverseTableName}id`]) {
        event.id = entity[`${databasePrefix}${eventsDataverseTableName}id`];
    }
    if (entity[`${databasePrefix}name`]) {
        event.name = entity[`${databasePrefix}name`];
    }
    if (entity[`${databasePrefix}startdate`]) {
        event.startDate = entity[`${databasePrefix}startdate`];
    }
    if (entity[`${databasePrefix}enddate`]) {
        event.endDate = entity[`${databasePrefix}enddate`];
    }
    if (entity[`${databasePrefix}readonly`]) {
        event.readOnly = entity[`${databasePrefix}readonly`];
    }
    if (entity[`${databasePrefix}timezone`]) {
        event.timeZone = entity[`${databasePrefix}timezone`];
    }
    if (entity[`${databasePrefix}draggable`]) {
        event.draggable = entity[`${databasePrefix}draggable`];
    }
    if (entity[`${databasePrefix}resizable`]) {
      if (entity[`${databasePrefix}resizable`] === "true") {
        event.resizable = true;
      } else if (entity[`${databasePrefix}resizable`] === "false") {
        event.resizable = false;
      } else {
        event.resizable = entity[`${databasePrefix}resizable`];
      }
    }
    if (entity[`${databasePrefix}children`]) {
        event.children = entity[`${databasePrefix}children`];
    }
    if (entity[`${databasePrefix}allday`]) {
        event.allDay = entity[`${databasePrefix}allday`];
    }
    if (entity[`${databasePrefix}duration`]) {
        event.duration = entity[`${databasePrefix}duration`];
    }
    if (entity[`${databasePrefix}durationunit`]) {
        event.durationUnit = entity[`${databasePrefix}durationunit`];
    }
    if (entity[`${databasePrefix}exceptiondates`]) {
        event.exceptionDates = JSON.parse(entity[`${databasePrefix}exceptiondates`]);
    }
    if (entity[`${databasePrefix}recurrencerule`]) {
        event.recurrenceRule = entity[`${databasePrefix}recurrencerule`];
    }
    if (entity[`${databasePrefix}cls`]) {
        event.cls = entity[`${databasePrefix}cls`];
    }
    if (entity[`${databasePrefix}eventcolor`]) {
        event.eventColor = entity[`${databasePrefix}eventcolor`];
    }
    if (entity[`${databasePrefix}eventstyle`]) {
        event.eventStyle = entity[`${databasePrefix}eventstyle`];
    }
    if (entity[`${databasePrefix}iconcls`]) {
        event.iconCls = entity[`${databasePrefix}iconcls`];
    }
    if (entity[`${databasePrefix}style`]) {
        event.style = entity[`${databasePrefix}style`];
    }
    return event;
  });
}

function removeResourcesDataColumnPrefixes(entities: any[]) {
  return entities.map((entity: ComponentFramework.WebApi.Entity) => {
    let resource: Partial<SchedulerResource> = {};
    if (entity[`${databasePrefix}${resourcesDataverseTableName}id`]) {
      resource.id = entity[`${databasePrefix}${resourcesDataverseTableName}id`];
    }
    if (entity[`${databasePrefix}name`]) {
      resource.name = entity[`${databasePrefix}name`];
    }
    if (entity[`${databasePrefix}eventcolor`]) {
      resource.eventColor = entity[`${databasePrefix}eventcolor`];
    }
    if (entity[`${databasePrefix}readonly`]) {
      resource.readOnly = entity[`${databasePrefix}readonly`];
    }
    return resource;
  });
}

function removeDependenciesDataColumnPrefixes(entities: any[]) {
  return entities.map((entity: ComponentFramework.WebApi.Entity) => {
    let dependency: Partial<SchedulerDependency> = {};
    if (entity[`${databasePrefix}${dependenciesDataverseTableName}id`]) {
      dependency.id =
        entity[`${databasePrefix}${dependenciesDataverseTableName}id`];
    }
    if (entity[`${databasePrefix}type`]) {
      dependency.type = entity[`${databasePrefix}type`];
    }
    if (entity[`${databasePrefix}cls`]) {
      dependency.cls = entity[`${databasePrefix}cls`];
    }
    if (entity[`${databasePrefix}lag`]) {
      dependency.lag = entity[`${databasePrefix}lag`];
    }
    if (entity[`${databasePrefix}lagunit`]) {
      dependency.lagUnit = entity[`${databasePrefix}lagunit`];
    }
    if (entity[`${databasePrefix}fromside`]) {
      dependency.fromSide = entity[`${databasePrefix}fromside`];
    }
    if (entity[`${databasePrefix}toside`]) {
      dependency.toSide = entity[`${databasePrefix}toside`];
    }
    if (entity[`${databasePrefix}from`]) {
      dependency.from =
        entity[`${databasePrefix}from`][
          `${databasePrefix}${eventsDataverseTableName}id`
        ];
    }
    if (entity[`${databasePrefix}to`]) {
      dependency.to =
        entity[`${databasePrefix}to`][
          `${databasePrefix}${eventsDataverseTableName}id`
        ];
    }
    return dependency;
  });
}

function removeAssignmentsDataColumnPrefixes(entities: any[]) {
  return entities.map((entity: ComponentFramework.WebApi.Entity) => {
    let assignment: Partial<SchedulerAssignment> = {};
    if (entity[`${databasePrefix}${assignmentsDataverseTableName}id`]) {
      assignment.id =
        entity[`${databasePrefix}${assignmentsDataverseTableName}id`];
    }
    if (entity[`${databasePrefix}eventid`]) {
      assignment.eventId =
        entity[`${databasePrefix}eventid`][
          `${databasePrefix}${eventsDataverseTableName}id`
        ];
    }
    if (entity[`${databasePrefix}resourceid`]) {
      assignment.resourceId =
        entity[`${databasePrefix}resourceid`][
          `${databasePrefix}${resourcesDataverseTableName}id`
        ];
    }
    return assignment;
  });
}

type NewAssignments = {
  // key is assignment phantomId, value is [eventId, resourceId]
  [key: string]: string[];
};

const BryntumScheduler: React.ComponentType<any> =
  OriginalBryntumScheduler as any;

const BryntumSchedulerComponent: FunctionComponent<
  IBryntumSchedulerComponentProps
> = (props) => {
  const [data, setData] = React.useState<{
    events: SchedulerEvent[];
    resources: SchedulerResource[];
    dependencies: SchedulerDependency[];
    assignments: SchedulerAssignment[];
  }>();

  const disableCreate = React.useRef(false);
  const newAssignmentResourceIdRef = useRef<NewAssignments>({});
  const scheduler = useRef<OriginalBryntumScheduler>(null);

  const fetchRecords = async () => {
    try {
      const eventsPromise = props?.context?.webAPI.retrieveMultipleRecords(
        `${databasePrefix}${eventsDataverseTableName}`
      );
      const resourcesPromise = props?.context?.webAPI.retrieveMultipleRecords(
        `${databasePrefix}${resourcesDataverseTableName}`
      );
      const dependenciesPromise:
        | Promise<ComponentFramework.WebApi.RetrieveMultipleResponse>
        | undefined = props?.context?.webAPI.retrieveMultipleRecords(
        `${databasePrefix}${dependenciesDataverseTableName}`,
        `?$select=${databasePrefix}type,${databasePrefix}lag,${databasePrefix}lagunit,${databasePrefix}fromside,${databasePrefix}toside,${databasePrefix}from,${databasePrefix}to&$expand=${databasePrefix}from($select=${databasePrefix}${eventsDataverseTableName}id),${databasePrefix}to($select=${databasePrefix}${eventsDataverseTableName}id)`
      );
      const assignmentsPromise:
        | Promise<ComponentFramework.WebApi.RetrieveMultipleResponse>
        | undefined = props?.context?.webAPI.retrieveMultipleRecords(
        `${databasePrefix}${assignmentsDataverseTableName}`,
        `?$select=${databasePrefix}eventid,${databasePrefix}resourceid&$expand=${databasePrefix}eventid($select=${databasePrefix}${eventsDataverseTableName}id),${databasePrefix}resourceid($select=${databasePrefix}${resourcesDataverseTableName}id)`
      );

      const [events, resources, dependencies, assignments] = await Promise.all([
        eventsPromise,
        resourcesPromise,
        dependenciesPromise,
        assignmentsPromise,
      ]);

      if (events && resources && dependencies && assignments) {
        setData({
          events: removeEventsDataColumnPrefixes(
            events.entities
          ) as SchedulerEvent[],
          resources: removeResourcesDataColumnPrefixes(
            resources.entities
          ) as SchedulerResource[],
          dependencies: removeDependenciesDataColumnPrefixes(
            dependencies.entities
          ) as SchedulerDependency[],
          assignments: removeAssignmentsDataColumnPrefixes(
            assignments.entities
          ) as SchedulerAssignment[],
        });
      }
    } catch (e) {
      if (e instanceof Error) {
        if (e.name === "PCFNonImplementedError") {
          console.log("PCFNonImplementedError: ", e.message);
          // You can add fallback data state for the development mode browser test harness
          // webAPI is not available in the test harness
        }
      }
      throw e;
    }
  };

  function onBeforeDragCreate() {
    disableCreate.current = true;
  }
  function onAfterDragCreate() {
    disableCreate.current = false;
  }

  const syncData = async ({ store, action, records }: SyncData) => {
    const storeId = store.id;

    if (storeId === "events") {
      if (action === "add") {
        for (let i = 0; i < records.length; i++) {
          return;
        }
      }
      if (action === "remove") {
        records.forEach((record) => {
          if (record.data.id.startsWith("_generated")) return;
          try {
            return props?.context?.webAPI
              .deleteRecord(
                `${databasePrefix}${eventsDataverseTableName}`,
                record.data.id
              )
              .then((res) => {
                console.log("deleteRecord: ", { res });
              });
          } catch (error) {
            console.error(error);
          }
        });
      }
      if (action === "update") {
        for (let i = 0; i < records.length; i++) {
          // create event
          if (records[i].data.id.startsWith("_generated")) {
            if (disableCreate.current === true) return;
            try {
              const newEvent = records[i] as SchedulerEvent;
              const newEventPhantomId = records[i].data.id;
              let insertData: Partial<SchedulerEventDataverse> = {};
              if (newEvent?.name) {
                insertData[`${databasePrefix}name`] = newEvent.name;
              }
              if (newEvent?.startDate) {
                insertData[`${databasePrefix}startdate`] = newEvent.startDate;
              }
              if (newEvent?.endDate) {
                insertData[`${databasePrefix}enddate`] = newEvent.endDate;
              }
              if (newEvent?.readOnly) {
                insertData[`${databasePrefix}readonly`] = Boolean(
                  newEvent.readOnly
                );
              }
              if (newEvent?.timeZone) {
                insertData[`${databasePrefix}timezone`] = newEvent.timeZone;
              }
              if (newEvent?.draggable) {
                insertData[`${databasePrefix}draggable`] = Boolean(
                  newEvent.draggable
                );
              }
              if (newEvent?.resizable) {
                insertData[
                  `${databasePrefix}resizable`
                ] = `${newEvent.resizable}`;
              }
              if (newEvent?.children) {
                insertData[`${databasePrefix}children`] = newEvent.children;
              }
              if (newEvent?.allDay) {
                insertData[`${databasePrefix}allday`] = Boolean(
                  newEvent.allDay
                );
              }
              if (newEvent?.duration) {
                insertData[`${databasePrefix}duration`] = Number(
                  newEvent.duration
                );
              }
              if (newEvent?.durationUnit) {
                insertData[`${databasePrefix}durationunit`] =
                  newEvent.durationUnit;
              }
              if (newEvent?.exceptionDates) {
                insertData[`${databasePrefix}exceptiondates`] = JSON.stringify(
                  newEvent.exceptionDates
                );
              }
              if (newEvent?.recurrenceRule) {
                insertData[`${databasePrefix}recurrencerule`] =
                  newEvent.recurrenceRule;
              }
              if (newEvent?.cls) {
                insertData[`${databasePrefix}cls`] = newEvent.cls;
              }
              if (newEvent?.eventColor) {
                insertData[`${databasePrefix}eventcolor`] = newEvent.eventColor;
              }
              if (newEvent?.eventStyle) {
                insertData[`${databasePrefix}eventstyle`] = newEvent.eventStyle;
              }
              if (newEvent?.iconCls) {
                insertData[`${databasePrefix}iconcls`] = newEvent.iconCls;
              }
              if (newEvent?.style) {
                insertData[`${databasePrefix}style`] = newEvent.style;
              }
              const createEventRecordRes =
                await props?.context?.webAPI.createRecord(
                  `${databasePrefix}${eventsDataverseTableName}`,
                  insertData
                );

              if (scheduler?.current?.instance && createEventRecordRes) {
                scheduler.current.instance.eventStore.applyChangeset({
                  updated: [
                    // Will set proper id for added event
                    {
                      $PhantomId: newEventPhantomId,
                      id: createEventRecordRes.id,
                    },
                  ],
                });
              }

              // create assignment for new event
              let newAssignmentData: string[] = [];

              for (const [key, value] of Object.entries(
                newAssignmentResourceIdRef.current
              )) {
                if (value[0] === newEventPhantomId) {
                  newAssignmentData = [key, ...value];
                }
              }

              let insertAssignmentsData: Partial<SchedulerAssignmentDataverse> =
                {};
              if (newAssignmentData.length !== 0 && createEventRecordRes) {
                insertAssignmentsData[
                  `${databasePrefix}eventid@odata.bind`
                ] = `/${databasePrefix}${eventsDataverseTableName}es(${createEventRecordRes?.id})`;

                const resourceId = newAssignmentData[2];
                insertAssignmentsData[
                  `${databasePrefix}resourceid@odata.bind`
                ] = `/${databasePrefix}${resourcesDataverseTableName}es(${resourceId})`;
                insertAssignmentsData[
                  `${databasePrefix}eventid@odata.bind`
                ] = `/${databasePrefix}${eventsDataverseTableName}es(${createEventRecordRes.id})`;
                const createAssignmentRecordRes =
                  await props?.context?.webAPI.createRecord(
                    `${databasePrefix}${assignmentsDataverseTableName}`,
                    insertAssignmentsData
                  );

                if (scheduler?.current?.instance && createAssignmentRecordRes) {
                  scheduler.current.instance.assignmentStore.applyChangeset({
                    updated: [
                      // Will set proper id for added assignment
                      {
                        $PhantomId: newAssignmentData[0],
                        id: createAssignmentRecordRes.id,
                      },
                    ],
                  });

                  let newObj = {
                    ...newAssignmentResourceIdRef.current,
                  };
                  delete newObj[newAssignmentData[0]];
                  newAssignmentResourceIdRef.current = newObj;
                }
              }
            } catch (error) {
              console.error(error);
            }

            // update event
          } else {
            try {
              if (records[i].data.id.startsWith("_generated")) continue;
              const event = records[i];
              let updateData: Partial<SchedulerEventDataverse> = {};
              if (event?.name) {
                updateData[`${databasePrefix}name`] = event.name;
              }
              if (event?.startDate) {
                updateData[`${databasePrefix}startdate`] = event.startDate;
              }
              if (event?.endDate) {
                updateData[`${databasePrefix}enddate`] = event.endDate;
              }
              if (event?.readOnly) {
                updateData[`${databasePrefix}readonly`] = Boolean(
                  event.readOnly
                );
              }
              if (event?.timeZone) {
                updateData[`${databasePrefix}timezone`] = event.timeZone;
              }
              if (event?.draggable) {
                updateData[`${databasePrefix}draggable`] = Boolean(
                  event.draggable
                );
              }
              if (event?.resizable) {
                updateData[`${databasePrefix}resizable`] = `${event.resizable}`;
              }
              if (event?.children) {
                updateData[`${databasePrefix}children`] = event.children;
              }
              if (event?.allDay) {
                updateData[`${databasePrefix}allday`] = Boolean(event.allDay);
              }
              if (event?.duration) {
                updateData[`${databasePrefix}duration`] = Number(
                  event.duration
                );
              }
              if (event?.durationUnit) {
                updateData[`${databasePrefix}durationunit`] =
                  event.durationUnit;
              }
              if (event?.exceptionDates) {
                updateData[`${databasePrefix}exceptiondates`] = JSON.stringify(
                  event.exceptionDates
                );
              }
              if (event?.recurrenceRule) {
                updateData[`${databasePrefix}recurrencerule`] =
                  event.recurrenceRule;
              }
              if (event?.cls) {
                updateData[`${databasePrefix}cls`] = event.cls;
              }
              if (event?.eventColor) {
                updateData[`${databasePrefix}eventcolor`] = event.eventColor;
              }
              if (event?.eventStyle) {
                updateData[`${databasePrefix}eventstyle`] = event.eventStyle;
              }
              if (event?.iconCls) {
                updateData[`${databasePrefix}iconcls`] = event.iconCls;
              }
              if (event?.style) {
                updateData[`${databasePrefix}style`] = event.style;
              }

              return props?.context?.webAPI
                .updateRecord(
                  `${databasePrefix}${eventsDataverseTableName}`,
                  event.id,
                  updateData
                )
                .then((res) => {
                  console.log("webAPI.updateRecord for events response: ", res);
                });
            } catch (error) {
              console.error(error);
            }
          }
        }
      }
    }
    if (storeId === "resources") {
      if (action === "add") {
        const resourcesIds = data?.resources.map((obj) => obj.id);
        for (let i = 0; i < records.length; i++) {
          const resourceExists =
            resourcesIds && resourcesIds.includes(records[i].data.id);
          if (resourceExists) return;
          try {
            const newResource = records[i] as SchedulerResource;
            let insertData: Partial<SchedulerResourceDataverse> = {};
            if (newResource?.name) {
              insertData[`${databasePrefix}name`] = newResource.name;
            }
            if (newResource?.eventColor) {
              insertData[`${databasePrefix}eventcolor`] =
                newResource.eventColor;
            }
            if (newResource?.readOnly) {
              insertData[`${databasePrefix}readonly`] = Boolean(
                newResource.readOnly
              );
            }

            return props?.context?.webAPI
              .createRecord(
                `${databasePrefix}${resourcesDataverseTableName}`,
                insertData
              )
              .then((res) => {
                if (scheduler?.current?.instance) {
                  scheduler.current.instance.resourceStore.applyChangeset({
                    updated: [
                      // Will set proper id for added resource
                      {
                        $PhantomId: newResource.id,
                        id: res.id,
                      },
                    ],
                  });
                }
              });
          } catch (error) {
            console.error(error);
          }
        }
      }

      if (action === "remove") {
        records.forEach((record) => {
          if (record.data.id.startsWith("_generated")) return;
          try {
            return props?.context?.webAPI
              .deleteRecord(
                `${databasePrefix}${resourcesDataverseTableName}`,
                record.data.id
              )
              .then((res) => {
                console.log("deleteRecord: ", { res });
              });
          } catch (error) {
            console.error(error);
          }
        });
      }

      if (action === "update") {
        for (let i = 0; i < records.length; i++) {
          try {
            if (records[i].data.id.startsWith("_generated")) return;
            const resource = records[i];
            let updateData: Partial<SchedulerResourceDataverse> = {};
            if (resource?.name) {
              updateData[`${databasePrefix}name`] = resource.name;
            }
            if (resource?.eventColor) {
              updateData[`${databasePrefix}eventcolor`] = resource.eventColor;
            }
            if (resource?.readOnly) {
              updateData[`${databasePrefix}readonly`] = Boolean(
                resource.readOnly
              );
            }

            return props?.context?.webAPI
              .updateRecord(
                `${databasePrefix}${resourcesDataverseTableName}`,
                resource.id,
                updateData
              )
              .then((res) => {
                console.log(
                  "webAPI.updateRecord for resources response: ",
                  res
                );
              });
          } catch (error) {
            console.error(error);
          }
        }
      }
    }

    if (storeId === "dependencies") {
      if (action === "add") {
        for (let i = 0; i < records.length; i++) {
          try {
            const newDependency = records[i] as SchedulerDependency;
            let insertData: Partial<SchedulerDependencyDataverse> = {};
            if (newDependency?.type) {
              insertData[`${databasePrefix}type`] = Number(newDependency.type);
            }
            if (newDependency?.cls) {
              insertData[`${databasePrefix}cls`] = newDependency.cls;
            }
            if (newDependency?.lag) {
              insertData[`${databasePrefix}lag`] = Number(newDependency.lag);
            }
            if (newDependency?.lagUnit) {
              insertData[`${databasePrefix}lagunit`] = newDependency.lagUnit;
            }
            if (newDependency?.fromSide) {
              insertData[`${databasePrefix}fromside`] = newDependency.fromSide;
            }
            if (newDependency?.toSide) {
              insertData[`${databasePrefix}toside`] = newDependency.toSide;
            }
            if (newDependency?.from) {
              insertData[
                `${databasePrefix}from@odata.bind`
              ] = `/${databasePrefix}${eventsDataverseTableName}es(${newDependency.from})`;
            }
            if (newDependency?.to) {
              insertData[
                `${databasePrefix}to@odata.bind`
              ] = `/${databasePrefix}${eventsDataverseTableName}es(${newDependency.to})`;
            }

            return props?.context?.webAPI
              .createRecord(
                `${databasePrefix}${dependenciesDataverseTableName}`,
                insertData
              )
              .then((res) => {
                if (scheduler?.current?.instance) {
                  scheduler.current.instance.dependencyStore.applyChangeset({
                    updated: [
                      // Will set proper id for added dependency
                      {
                        $PhantomId: newDependency.id,
                        id: res.id,
                      },
                    ],
                  });
                }
              });
          } catch (error) {
            console.error(error);
          }
        }
      }
      if (action === "remove") {
        records.forEach((record) => {
          if (record.data.id.startsWith("_generated")) return;
          try {
            return props?.context?.webAPI
              .deleteRecord(
                `${databasePrefix}${dependenciesDataverseTableName}`,
                record.data.id
              )
              .then((res) => {
                console.log("deleteRecord: ", { res });
              });
          } catch (error) {
            console.error(error);
          }
        });
      }
      if (action === "update") {
        for (let i = 0; i < records.length; i++) {
          try {
            if (records[i].data.id.startsWith("_generated")) return;
            const dependency = records[i];
            let updateData: Partial<SchedulerDependencyDataverse> = {};
            if (dependency?.type) {
              updateData[`${databasePrefix}type`] = Number(dependency.type);
            }
            if (dependency?.cls) {
              updateData[`${databasePrefix}cls`] = dependency.cls;
            }
            if (dependency?.lag) {
              updateData[`${databasePrefix}lag`] = Number(dependency.lag);
            }
            if (dependency?.lagUnit) {
              updateData[`${databasePrefix}lagunit`] = dependency.lagUnit;
            }
            if (dependency?.fromSide) {
              updateData[`${databasePrefix}fromside`] = dependency.fromSide;
            }
            if (dependency?.toSide) {
              updateData[`${databasePrefix}toside`] = dependency.toSide;
            }
            if (dependency?.from) {
              updateData[
                `${databasePrefix}from@odata.bind`
              ] = `/${databasePrefix}${eventsDataverseTableName}es(${dependency.from})`;
            }
            if (dependency?.to) {
              updateData[
                `${databasePrefix}to@odata.bind`
              ] = `/${databasePrefix}${eventsDataverseTableName}es(${dependency.to})`;
            }

            return props?.context?.webAPI
              .updateRecord(
                `${databasePrefix}${dependenciesDataverseTableName}`,
                dependency.id,
                updateData
              )
              .then((res) => {
                console.log(
                  "webAPI.updateRecord for dependencies response: ",
                  res
                );
              });
          } catch (error) {
            console.error(error);
          }
        }
      }
    }
    if (storeId === "assignments") {
      if (action === "add") {
        for (let i = 0; i < records.length; i++) {
          const newAssignment = records[i] as SchedulerAssignment;

          // add assignment for new event
          if (newAssignment.eventId.startsWith("_generated")) {
            newAssignmentResourceIdRef.current[newAssignment.id] = [
              newAssignment.eventId,
              newAssignment.resourceId,
            ];

            // add assignment for existing event
          } else {
            try {
              let insertData: Partial<SchedulerAssignmentDataverse> = {};
              if (newAssignment?.eventId) {
                insertData[
                  `${databasePrefix}eventid@odata.bind`
                ] = `/${databasePrefix}${eventsDataverseTableName}es(${newAssignment.eventId})`;
              }
              if (newAssignment?.resourceId) {
                insertData[
                  `${databasePrefix}resourceid@odata.bind`
                ] = `/${databasePrefix}${resourcesDataverseTableName}es(${newAssignment.resourceId})`;
              }
              return props?.context?.webAPI
                .createRecord(
                  `${databasePrefix}${assignmentsDataverseTableName}`,
                  insertData
                )
                .then((res) => {
                  if (scheduler?.current?.instance) {
                    scheduler.current.instance.assignmentStore.applyChangeset({
                      updated: [
                        // Will set proper id for added assignment
                        {
                          $PhantomId: newAssignment.id,
                          id: res.id,
                        },
                      ],
                    });
                  }
                });
            } catch (error) {
              console.error(error);
            }
          }
        }
      }

      if (action === "remove") {
        records.forEach((record) => {
          if (record.data.id.startsWith("_generated")) return;
          try {
            return props?.context?.webAPI
              .deleteRecord(
                `${databasePrefix}${assignmentsDataverseTableName}`,
                record.data.id
              )
              .then((res) => {
                console.log("deleteRecord: ", { res });
              });
          } catch (error) {
            console.error(error);
          }
        });
      }
      if (action === "update") {
        for (let i = 0; i < records.length; i++) {
          try {
            if (records[i].data.id.startsWith("_generated")) return;
            const assignment = records[i];
            let updateData: Partial<SchedulerAssignmentDataverse> = {};
            if (assignment?.eventId) {
              updateData[
                `${databasePrefix}eventid@odata.bind`
              ] = `/${databasePrefix}${eventsDataverseTableName}es(${assignment.eventId})`;
            }
            if (assignment?.resourceId) {
              updateData[
                `${databasePrefix}resourceid@odata.bind`
              ] = `/${databasePrefix}${resourcesDataverseTableName}es(${assignment.resourceId})`;
            }

            return props?.context?.webAPI
              .updateRecord(
                `${databasePrefix}${assignmentsDataverseTableName}`,
                assignment.id,
                updateData
              )
              .then((res) => {
                console.log(
                  "webAPI.updateRecord for assignments response: ",
                  res
                );
              });
          } catch (error) {
            console.error(error);
          }
        }
      }
    }
  };

  React.useEffect(() => {
    fetchRecords();
  }, []);

  return data ? (
    <BryntumScheduler
      ref={scheduler}
      events={data.events}
      resources={data.resources}
      dependencies={data.dependencies}
      assignments={data.assignments}
      onDataChange={syncData}
      onBeforeDragCreate={onBeforeDragCreate}
      onAfterDragCreate={onAfterDragCreate}
      {...schedulerConfig}
    />
  ) : (
    <LoadingSpinner />
  );
};

export default BryntumSchedulerComponent;
