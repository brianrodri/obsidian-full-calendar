import ical from "ical.js";
import { OFCEvent, validateEvent } from "../../types";
import { DateTime } from "luxon";
import { rrulestr } from "rrule";

function getDate(t: ical.Time, tz: ical.Timezone): string {
    return DateTime.fromSeconds(t.toUnixTime()).setZone(tz.tzid).toISODate();
}

function getTime(t: ical.Time, tz: ical.Timezone): string {
    if (t.isDate) {
        return "00:00";
    }
    return DateTime.fromSeconds(t.toUnixTime()).setZone(tz.tzid).toISOTime({
        includeOffset: false,
        includePrefix: false,
        suppressMilliseconds: true,
        suppressSeconds: true,
    });
}

function extractEventUrl(iCalEvent: ical.Event): string {
    let urlProp = iCalEvent.component.getFirstProperty("url");
    return urlProp ? urlProp.getFirstValue() : "";
}

// NOTE: Google Calendar uses this ICAL property to specify a timezone for its events.
const GCAL_TZID_PROP_NAME = "x-wr-timezone";

function extractTimezone(calComponent: ical.Component): ical.Timezone {
    if (calComponent.hasProperty(GCAL_TZID_PROP_NAME)) {
        const tzid = calComponent.getFirstPropertyValue(GCAL_TZID_PROP_NAME);
        for (const component of calComponent.getAllSubcomponents("vtimezone")) {
            if (component.getFirstPropertyValue("tzid") === tzid) {
                return new ical.Timezone({ component });
            }
        }
        return new ical.Timezone({ tzid, component: "" });
    } else {
        const component = calComponent.getFirstSubcomponent("vtimezone");
        return component
            ? new ical.Timezone({ component })
            : ical.Timezone.utcTimezone;
    }
}

function specifiesEnd(iCalEvent: ical.Event) {
    return (
        Boolean(iCalEvent.component.getFirstProperty("dtend")) ||
        Boolean(iCalEvent.component.getFirstProperty("duration"))
    );
}

function icsToOFC(input: ical.Event, tz: ical.Timezone): OFCEvent {
    const allDay = input.startDate.isDate;
    if (allDay) {
        tz = ical.Timezone.utcTimezone;
    }

    if (input.isRecurring()) {
        const rrule = rrulestr(
            input.component.getFirstProperty("rrule").getFirstValue().toString()
        );
        const exdates = input.component
            .getAllProperties("exdate")
            .map((exdateProp) => {
                const exdate = exdateProp.getFirstValue();
                // NOTE: We only store the date from an exdate and recreate the full datetime exdate later,
                // so recurring events with exclusions that happen more than once per day are not supported.
                return getDate(exdate, tz);
            });

        return {
            type: "rrule",
            title: input.summary,
            id: `ics::${input.uid}::${getDate(input.startDate, tz)}::recurring`,
            rrule: rrule.toString(),
            skipDates: exdates,
            startDate: getDate(input.startDate, tz),
            ...(allDay
                ? { allDay: true }
                : {
                      allDay: false,
                      startTime: getTime(input.startDate, tz),
                      endTime: input.endDate
                          ? getTime(input.endDate, tz)
                          : null,
                  }),
        };
    } else {
        const date = getDate(input.startDate, tz);
        const endDate = specifiesEnd(input) ? getDate(input.endDate, tz) : null;
        return {
            type: "single",
            id: `ics::${input.uid}::${date}::single`,
            title: input.summary,
            date,
            endDate: date !== endDate ? endDate : null,
            ...(allDay
                ? { allDay: true }
                : {
                      allDay: false,
                      startTime: getTime(input.startDate, tz),
                      endTime: getTime(input.endDate, tz),
                  }),
        };
    }
}

export function getEventsFromICS(text: string): OFCEvent[] {
    const jCalData = ical.parse(text);
    const component = new ical.Component(jCalData);
    const tz = extractTimezone(component);

    const events: ical.Event[] = component
        .getAllSubcomponents("vevent")
        .map((vevent) => new ical.Event(vevent))
        .filter((evt) => {
            evt.iterator;
            try {
                evt.startDate.toJSDate();
                evt.endDate.toJSDate();
                return true;
            } catch (err) {
                // skipping events with invalid time
                return false;
            }
        });

    // Events with RECURRENCE-ID will have duplicated UIDs.
    // We need to modify the base event to exclude those recurrence exceptions.
    const baseEvents = Object.fromEntries(
        events
            .filter((e) => e.recurrenceId === null)
            .map((e) => [e.uid, icsToOFC(e, tz)])
    );

    const recurrenceExceptions = events
        .filter((e) => e.recurrenceId !== null)
        .map((e): [string, OFCEvent] => [e.uid, icsToOFC(e, tz)]);

    for (const [uid, event] of recurrenceExceptions) {
        const baseEvent = baseEvents[uid];
        if (!baseEvent) {
            continue;
        }

        if (baseEvent.type !== "rrule" || event.type !== "single") {
            console.warn(
                "Recurrence exception was recurring or base event was not recurring",
                { baseEvent, recurrenceException: event }
            );
            continue;
        }
        baseEvent.skipDates.push(event.date);
    }

    const allEvents = Object.values(baseEvents).concat(
        recurrenceExceptions.map((e) => e[1])
    );

    return allEvents.map(validateEvent).flatMap((e) => (e ? [e] : []));
}
