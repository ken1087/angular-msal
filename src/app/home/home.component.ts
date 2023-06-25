import { Component, OnInit } from '@angular/core';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType, InteractionStatus } from '@azure/msal-browser';
import { Client } from '@microsoft/microsoft-graph-client';
import { endOfWeek, startOfWeek } from 'date-fns';
import { zonedTimeToUtc } from 'date-fns-tz';
import { filter } from 'rxjs/operators';
import { findIana } from 'windows-iana';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {
  loginDisplay = false;

  constructor(
    private authService: MsalService, 
    private msalBroadcastService: MsalBroadcastService
    
  ) { }

  ngOnInit(): void {
    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe((result: EventMessage) => {
        console.log(result);
        const payload = result.payload as AuthenticationResult;
        this.authService.instance.setActiveAccount(payload.account);
      });
    
    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None)
      )
      .subscribe(() => {
        this.setLoginDisplay();
      })
    
  }
  
  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
  }

  async onClickGetCalander() {
    // Convert the user's timezone to IANA format
    const ianaName = findIana('UTC');
    const timeZone = 'UTC';

    // Get midnight on the start of the current week in the user's timezone,
    // but in UTC. For example, for Pacific Standard Time, the time value would be
    // 07:00:00Z
    const now = new Date();
    const weekStart = zonedTimeToUtc(startOfWeek(now), timeZone);
    const weekEnd = zonedTimeToUtc(endOfWeek(now), timeZone);
    try {
      
      console.log(sessionStorage.getItem('accessToken'));

      const graphClient = Client.init({
        // Initialize the Graph client with an auth
        // provider that requests the token from the
        // auth service
        authProvider: async(done) => {
          done(null, sessionStorage.getItem('accessToken'));
        }
      });

      const graphUser: MicrosoftGraph.User = await graphClient
        .api('/me')
        .select('displayName')
        .get();
      console.log(graphUser);
      // GET /me/calendarview?startDateTime=''&endDateTime=''
      // &$select=subject,organizer,start,end
      // &$orderby=start/dateTime
      // &$top=50
      const result =  await graphClient
        .api('/me/calendarview')
        .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({
          startDateTime: weekStart.toISOString(),
          endDateTime: weekEnd.toISOString()
        })
        .select('start,end')
        .orderby('start/dateTime')
        .top(50)
        .get();

      console.log(result);

    } catch (error) {
      console.log(error);
    }

  }

}
