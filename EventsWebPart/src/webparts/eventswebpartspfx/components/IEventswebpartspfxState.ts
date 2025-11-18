export interface IEvent {
    Id: number;
    Title: string;
    EventDate: string;
    EndDate: string;
    Location?: string;
  }
  
  export interface IEventswebpartspfxState {
    events: IEvent[];
  }