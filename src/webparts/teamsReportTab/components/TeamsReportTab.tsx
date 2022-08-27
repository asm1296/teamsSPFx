import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as pnp from "sp-pnp-js";
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
/* import { SPHttpClient } from '@microsoft/sp-http';
import { spODataEntityArray, Item, IItem } from "@pnp/sp/presets/all"; */
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/items/list";
import "@pnp/sp/fields";
import "@pnp/sp/views";

import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";
import styles from './TeamsReportTab.module.scss';
import { Provider, teamsTheme, Grid, Input, Text, Button,TextArea } from '@fluentui/react-northstar';

import { ITeamsReportTabProps } from './ITeamsReportTabProps';
import { ITeamsReportTabState} from './ITeamsReportTabState';

export default class TeamsReportTab extends React.Component<ITeamsReportTabProps,ITeamsReportTabState>{
  constructor(props:ITeamsReportTabProps){
    super(props);
    this.state={
      Date : new Date(),
      ticketHandled : 0,
      userName: "",
      incidentResolved : 0,
      requestResolved : 0,
      ticketRouted : 0,
      ticketOnHold : 0,
      ticketInProg : 0,
      additionalTask : ""
        };
  }

   public handleEventChange = (event:React.ChangeEvent<HTMLInputElement>)=>{
     event.persist();
     this.setState((prevState)=>({
      ...prevState,
      [event.target.name as keyof ITeamsReportTabState] : event.target.value
     }));

  }

    
   public currentUser(){
      pnp.sp.web.currentUser.get().then(user=>{
        this.setState({
          userName : user.Title
        });
      }).catch(err=>{
        console.log("Getting error in querying User Details");
      });
    }

    public async createListWithFields():Promise<void>{
      var dateobj=this.state.Date;
       let year = dateobj.getFullYear();
       let month= ("0" + (dateobj.getMonth()+1)).slice(-2);
       let date = ("0" + dateobj.getDate()).slice(-2);
       let converted_date = year + "-" + month + "-" + date; 
      let listEnsureResult = sp.web.lists.ensure("O365-ReportData_"+converted_date);
      const dateSchema = `<Field ID="{44A1174E-8C63-4099-A354-0525C6902C35}"
      Name="Date"
      DisplayName="Date"
      Type="DateTime"
      Format="DateOnly"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const userNameSchema = `<Field ID="{6C4EE2B2-F231-4B05-A613-70F5551B946B}"
      Name="userName"
      DisplayName="userName"
      Type="Text"
      Required="TRUE"
      Group="TeamsReportTab Columns">
</Field>`;
      const ticketHandledSchema = `<Field ID="{AD4F9AF9-80CC-4A27-B0BD-96E9558BBE77}"
      Name="ticketHandled"
      DisplayName="ticketHandled"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const incidentResolvedSchema = `<Field ID="{A7DA941D-BA8C-47E6-A4C5-ACF65B01A188}"
      Name="incidentResolved"
      DisplayName="incidentResolved"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const requestResolvedSchema = `<Field ID="{1FE3F9B2-8557-4387-8810-C3950B5D577A}"
      Name="requestResolved"
      DisplayName="requestResolved"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const ticketRoutedSchema = `<Field ID="{BF12FCB0-48EA-4783-BEB6-EDD0CC7930DE}"
      Name="ticketRouted"
      DisplayName="ticketRouted"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const ticketOnHoldSchema = `<Field ID="{A82B1906-8CFE-4CDF-A35F-FA9E0AD2A06B}"
      Name="ticketOnHold"
      DisplayName="ticketOnHold"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;
      const ticketInProgressSchema = `<Field ID="{30EAAEC3-8E66-4350-9FC2-79A13410DDCB}"
      Name="ticketInProg"
      DisplayName="ticketInProg"
      Type="Number"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`; 
      const additionalTaskSchema = `<Field ID="{B0D20685-C219-4044-ABB2-B43DEEAFDFDB}"
      Name="additionalTask"
      DisplayName="additionalTask"
      Type="Note"
      Required="FALSE"
      Group="TeamsReportTab Columns">
</Field>`;

      const allSchema = [dateSchema, userNameSchema, ticketHandledSchema, incidentResolvedSchema, requestResolvedSchema, ticketRoutedSchema, ticketOnHoldSchema, ticketInProgressSchema, additionalTaskSchema];
      const fields = ['Date','userName','ticketHandled','incidentResolved','requestResolved','ticketRouted','ticketOnHold','ticketInProg','additionalTask' ];
      let listd = sp.web.lists.getByTitle("O365-ReportData_"+converted_date);
      let viewd = listd.defaultView;
      if ((await listEnsureResult).created){
        /* sp.web.lists.getByTitle("O365-ReportData_"+converted_date).fields.createFieldAsXml(dateSchema);*/
        const batchAddField = sp.web.createBatch();
        const batchView = sp.web.createBatch();
        allSchema.forEach(schema =>{
          listd.fields.inBatch(batchAddField).createFieldAsXml(schema);
        });
        batchAddField.execute().then(()=>{
          viewd.fields.inBatch(batchView).removeAll();
          fields.forEach(field=>{
            viewd.fields.inBatch(batchView).add(field);
          });
          batchView.execute().then(()=>{
            console.log('Fields and Views are created');
          }).catch((err)=>{
            console.log(err);
          });
        }).catch((err)=>{
          console.log(err);
        });


      }
      else {
        console.log("List and View  is already created");
      }
    }

    public componentDidMount(){
      this.currentUser();  
      this.createListWithFields(); 
    }

    public changeDateHandler =date=>{
    this.setState({
      Date:date
    },()=>(
      this.createListWithFields()
      ));
  }
  
  public async addItems():Promise<void>{
    var dateobj=this.state.Date;
       let year = dateobj.getFullYear();
       let month= ("0" + (dateobj.getMonth()+1)).slice(-2);
       let date = ("0" + dateobj.getDate()).slice(-2);
       let converted_date = year + "-" + month + "-" + date; 
      let newitem : ITeamsReportTabState={Date:this.state.Date,userName : this.state.userName,ticketHandled : this.state.ticketHandled ,incidentResolved : this.state.incidentResolved , requestResolved : this.state.requestResolved, ticketRouted : this.state.ticketRouted, ticketOnHold : this.state.ticketOnHold, ticketInProg : this.state.ticketInProg, additionalTask : this.state.additionalTask };
    sp.web.lists.getByTitle("O365-ReportData_"+converted_date).items.add(newitem);
    ReactDom.render(<h3>Report Updated- Created now</h3>,document.getElementById("updateStatus")); 
}

/* public async GetItems():Promise<any>{
var dateobj=this.state.Date
       let year = dateobj.getFullYear();
       let month= ("0" + (dateobj.getMonth()+1)).slice(-2);
       let date = ("0" + dateobj.getDate()).slice(-2);
       let converted_date = year + "-" + month + "-" + date; 
const results = sp.web.lists.getByTitle("O365-ReportData_"+converted_date).items.get();
console.log(results);
let modresults = (await results).map(result=><table><tr><td>{result.userName}</td><td>{result.ticketHandled}</td><td>{result.additionalTask}</td></tr></table>);
ReactDom.render(<p>{modresults}</p>,document.getElementById("showListItems"));
} */

  public render(): React.ReactElement<ITeamsReportTabProps> {
    
    return (
    <Provider theme={teamsTheme}>
      <div>
        <div>
          <div>
            <div>
              <section className={styles.arrangeHeader} style={{backgroundColor : "#6264A7"}}>
              <span className={ styles.title }>{this.props.title}</span>
              <p className={ styles.subTitle }>{this.props.subtitle}</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              </section>
              <section className={styles.arrangeMain}>
              <div className={styles.dateField}>         
              <label> Date: </label>
              <DatePicker selected={this.state.Date} onChange={this.changeDateHandler}></DatePicker>
              </div>
               <Grid styles={{gridTemplateColumns: '1fr 2fr',msGridColumns: '1fr 2fr',gap : '10px',textAlign:'center'}}>
                <Text styles={{placeSelf : 'center center'}}>Ticket Handled: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="ticketHandled" value={this.state.ticketHandled} onChange = {this.handleEventChange}></Input>
              <Text styles={{placeSelf : 'center center'}}>Incident Resolved: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="incidentResolved" value={this.state.incidentResolved} onChange = {this.handleEventChange}></Input>
              <Text styles={{placeSelf : 'center center'}}>Request Resolved: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="requestResolved" value={this.state.requestResolved} onChange = {this.handleEventChange}></Input>
              <Text styles={{placeSelf : 'center center'}}>Ticket Routed: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="ticketRouted" value={this.state.ticketRouted} onChange = {this.handleEventChange}></Input>
              <Text styles={{placeSelf : 'center center'}}>Ticket OnHold: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="ticketOnHold" value={this.state.ticketOnHold} onChange = {this.handleEventChange}></Input>
             <Text styles={{placeSelf : 'center center'}}>Ticket Inprogress: </Text>
                <Input inverted styles={{placeSelf : 'center start'}} type="text" name="ticketInProg" value={this.state.ticketInProg} onChange = {this.handleEventChange}></Input>
             <Text styles={{placeSelf : 'center center'}}>Additional Task: </Text>
                <TextArea inverted styles={{placeSelf : 'center start'}} name="additionalTask" value={this.state.additionalTask} onChange = {this.handleEventChange}></TextArea>
                <br />
             <Button primary styles={{placeSelf : 'center start'}} onClick={this.addItems.bind(this)}>Submit</Button>
             {/* <Button primary styles={{placeSelf : 'center start'}} onClick={this.GetItems.bind(this)}>Get Items</Button> */}
             <div id="updateStatus">
              </div>
                </Grid>
              </section>
                    
            </div>
          </div>
        </div>
      </div>
      </Provider>
    );
  }
}
