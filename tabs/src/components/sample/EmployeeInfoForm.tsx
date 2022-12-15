// import { useState } from "react";
import { Button, Input } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { useState } from "react";
import '../tabStyles.css';

const EmployeeInfoForm = (props: any) => {
  
  const initialState = {
    name: '',
    role: '',
    topic: ''
  }

  const [state, setState] = useState(initialState);

  const submitEmployeeData = () => {
    const empData = {
      Name: state.name,
      Role: state.role,
      Topic: state.topic
    }

    microsoftTeams.dialog.submit(empData);

  }

  return (
    <div className={`ms-Grid`}>
      <div className={`ms-Grid-row`} 
      style={{
        background: "radial-gradient(circle, rgba(238,174,202,1) 0%, rgba(148,187,233,1) 100%)", 
        padding: "5rem 2rem 10rem 2rem"
      }}>
        <div className={`ms-Grid-col ms-sm12`}>
          <h1>Registration Form</h1>
        </div>
        <div className={`ms-Grid-row`}>
          <div className={`ms-Grid-col ms-sm3`}>
            <Input type="text" label="Participant Name" onChange={(e, ev)=>{setState({...state, name: ev?.value ? ev.value : '' })}}></Input>
          </div>
          <div className={`ms-Grid-col ms-sm3`}>
            <Input type="text" label="Participant Role" onChange={(e, ev)=>{setState({...state, role: ev?.value ? ev.value : '' })}}></Input>
          </div>
          <div className={`ms-Grid-col ms-sm3`}>
            <Input type="text" label="Participant Topic" onChange={(e, ev)=>{setState({...state, topic: ev?.value ? ev.value : '' })}}></Input>
          </div>
          <div className={`ms-Grid-col ms-sm3`} style={{marginTop: "1rem", textAlign: "right"}}>
            <Button primary onClick={()=>submitEmployeeData()}>Submit</Button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default EmployeeInfoForm;
