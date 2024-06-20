import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { sp } from "@pnp/sp/presets/all";
import { IAppointmentFormProps } from './IAppointmentFormProps';
import { DefaultButton } from '@fluentui/react';
// import {ILoginProps} from './ILoginProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import 'bootstrap/dist/css/bootstrap.min.css';
// import {peoplepic}
// import { DefaultButton } from '@fluentui/react';
// import styles from './AppointmentForm.module.scss';

SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js').catch((error: any) => {
  console.log(error);
});





export default class AppointmentForm extends React.Component<IAppointmentFormProps, any> {
  [x: string]: any;
  constructor(props: IAppointmentFormProps) {
    super(props);
    this.state = {
      items: [],
      selectedItemID: null
    };
    sp.setup({
      spfxContext: this.props.context as any
    })

    this.savedata = this.savedata.bind(this);
    this.updatedata = this.updatedata.bind(this);
    this.deletedata = this.deletedata.bind(this);
    this.readdata = this.readdata.bind(this);
    this.editItem = this.editItem.bind(this);
  }

  public async componentDidMount() {
    this.readdata();
  }


  public render(): React.ReactElement<IAppointmentFormProps> {
    return (
      <>
        <div className="form-container">
          <form onSubmit={this.savedata}>
            <div className='row'>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Name:</label>
                <input className='form-control' type="text" id="name" name="name" required />
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Phone:</label>
                <input className='form-control' type="number" id="phone" name="phone" required />
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Email:</label>
                <input className='form-control' type="email" id="email" name="email" required />
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Gender:</label>
                <select className='form-control' id="gender" name="gender" required>
                  <option value="">Select</option>
                  <option value="Male">Male</option>
                  <option value="Female">Female</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Date of Birth:</label>
                <input className='form-control' type="date" id="dob" name="Date of Birth" required />
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Message:</label>
                <textarea className='form-control' id="message" name="message" required />
              </div>
              <div className='col-sm-4 mb-2'>
                <label className='col-sm-4'>Assignee:</label>
                <div className='form-input' id ="assigneePerson" >

                  <PeoplePicker
                    peoplePickerCntrlclassName='form-input'
                    context={this.props.spfxContext as any}
                    personSelectionLimit={1}
                    showtooltip={false}
                    required={false}
                    disabled={false}
                    showHiddenInUI={false}
                    ensureUser={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={500}
                    onChange={(people: any) => this.setState({ selectedPeople: people })}
                  />
                </div>
              </div>
            </div>
            <div className='col-sm-4 mb-2'>
              <input className='btn btn-primary col-sm-2' type="submit" value="Submit" />
              <input className='btn btn-secondary col-sm-2' type="button" onClick={this.updatedata} value="Update" />
              <input className='btn btn-danger col-sm-2' type="button" onClick={() => this.deletedata()} value="Delete" />
              <input className='btn btn-info col-sm-2 ml-2' type="button" onClick={this.readdata} value="Read Data" />
            </div>
          </form>
          <div>
            <h2>Fetched Items</h2>
            <table className='table table-bordered'>
              <tr>
                <th>Sr. No.</th>
                <th>Name</th>
                <th>Phone</th>
                <th>Email</th>
                <th>Message</th>
                <th>Date of Birth</th>
                <th>Gender</th>
                <th>Edit/Delete</th>
              </tr>
              <tbody>
                {this.state.items.map((item: any, index: number) => (
                  <tr key={index}>
                    <td>{index + 1}</td>
                    <td>{item.Name}</td>
                    <td>{item.MobileNumber}</td>
                    <td>{item.Email}</td>
                    <td>{item.Message}</td>
                    <td>{item.DateofBirth}</td>
                    <td>{item.Gender}</td>
                    <td>
                      <DefaultButton className='btn btn-primary m-1' text="Edit" onClick={() => this.editItem(item)} />
                      <DefaultButton className='btn btn-danger m-1' text="Delete" onClick={() => this.deletedata(item.Id)} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </>
    );
  }

  public async savedata(event: React.FormEvent) {
    event.preventDefault();

    let obj: any = {
      Name: (document.getElementById("name") as HTMLInputElement).value,
      MobileNumber: (document.getElementById("phone") as HTMLInputElement).value,
      Email: (document.getElementById("email") as HTMLInputElement).value,
      Gender: (document.getElementById("gender") as HTMLInputElement).value,
      DateofBirth: (document.getElementById("dob") as HTMLInputElement).value,
      Message: (document.getElementById("message") as HTMLInputElement).value,
      PeopleId: {
        results: this.state.selectedPeople.map((person: any) => person.id)
      }
    };

    try {
      let result = await sp.web.lists.getByTitle("AppointmentForm").items.add(obj);
      console.log("Data saved successfully.", result);
      alert("Data saved successfully.")
      await this.readdata().catch(e => (e));
      this.resetForm();
    } catch (error) {
      console.error("Error saving data", error);
    }
  }
  private resetForm() {
    (document.getElementById("name") as HTMLInputElement).value = '';
    (document.getElementById("phone") as HTMLInputElement).value = '';
    (document.getElementById("email") as HTMLInputElement).value = '';
    (document.getElementById("gender") as HTMLInputElement).value = '';
    (document.getElementById("dob") as HTMLInputElement).value = '';
    (document.getElementById("message") as HTMLInputElement).value = '';
    this.setState({ selectedPeople: [] });
  }

  public async updatedata() {
    if (this.state.selectedItemID) {
      let obj: any = {
        Name: (document.getElementById("name") as HTMLInputElement).value,
        MobileNumber: (document.getElementById("phone") as HTMLInputElement).value,
        Email: (document.getElementById("email") as HTMLInputElement).value,
        Gender: (document.getElementById("gender") as HTMLInputElement).value,
        DateofBirth: (document.getElementById("dob") as HTMLInputElement).value,
        Message: (document.getElementById("message") as HTMLInputElement).value
      };

      try {
        let result = await sp.web.lists.getByTitle("AppointmentForm").items.getById(this.state.selectedItemID).update(obj);
        console.log("Data updated successfully", result);
        alert("Data updated successfully")
        await this.readdata().catch(e => (e));
        this.resetForm();
      } catch (error) {
        console.error("Error updating data", error);
      }
    } else {
      console.error("No item selected for update");
    }
  }

  public async deletedata(id?: number) {
    const itemId = id || this.state.selectedItemID;
    if (itemId) {
      const confirmDeletion = confirm("Are you sure you want to delete this item?");
      if (confirmDeletion) {
        try {
          let result = await sp.web.lists.getByTitle("AppointmentForm").items.getById(itemId).delete();
          console.log("Data deleted successfully", result);
          await this.readdata();
          alert("Data deleted successfully.")
          this.setState({ selectedItemID: null });
        } catch (error) {
          console.error("Error deleting data", error);
        }
      } else {
        console.log("Deletion cancelled by user.");
      }
    } else {
      console.error("No item selected for deletion");
    }
  }


  public async readdata() {
    try {
      let items: any[] = await sp.web.lists.getByTitle("AppointmentForm").items.getAll();
      this.setState({ items });
      console.log("Data read successfully");
    } catch (error) {
      console.error("Error reading data", error);
    }
  }

  private editItem(item: any) {
    this.setState({ selectedItemID: item.Id });
    (document.getElementById("name") as HTMLInputElement).value = item.Name;
    (document.getElementById("phone") as HTMLInputElement).value = item.MobileNumber;
    (document.getElementById("email") as HTMLInputElement).value = item.Email;
    (document.getElementById("gender") as HTMLInputElement).value = item.Gender;
    (document.getElementById("dob") as HTMLInputElement).value = item.DateofBirth;
    (document.getElementById("message") as HTMLInputElement).value = item.Message;
  }
}
// import * as React from 'react';
// import { SPComponentLoader } from '@microsoft/sp-loader';
// import { IAppointmentFormProps } from './IAppointmentFormProps';
// import { DefaultButton } from '@fluentui/react';
// import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// // import 'bootstrap/dist/css/bootstrap.min.css';
// // import styles from './AppointmentForm.module.scss';

// SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js').catch((error: any) => {
//   console.log(error);
// });

// export default class AppointmentForm extends React.Component<IAppointmentFormProps, any> {
//   [x: string]: any;
//   constructor(props: IAppointmentFormProps) {
//     super(props);
//     this.state = {
//       items: [],
//       selectedItemID: null,
//       selectedPeople: []
//     };

//     this.savedata = this.savedata.bind(this);
//     this.updatedata = this.updatedata.bind(this);
//     this.deletedata = this.deletedata.bind(this);
//     this.readdata = this.readdata.bind(this);
//     this.editItem = this.editItem.bind(this);
//   }

//   public async componentDidMount() {
//     this.readdata();
//     //console.log("componentdidmount")
//   }

//   public render(): React.ReactElement<IAppointmentFormProps> {
//     //console.log("render")
//     return (
//       <>
//         <div className="form-container">
//           <form onSubmit={this.savedata}>
//             <div className='row'>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Nasdme:</label>
//                 <input className='form-control' type="text" id="name" name="name" required />
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Phone:</label>
//                 <input className='form-control' type="number" id="phone" name="phone" required />
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Email:</label>
//                 <input className='form-control' type="email" id="email" name="email" required />
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Gender:</label>
//                 <select className='form-control' id="gender" name="gender" required>
//                   <option value="">Select</option>
//                   <option value="Male">Male</option>
//                   <option value="Female">Female</option>
//                   <option value="Other">Other</option>
//                 </select>
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Date of Birth:</label>
//                 <input className='form-control' type="date" id="dob" name="Date of Birth" required />
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Message:</label>
//                 <textarea className='form-control' id="message" name="message" required />
//               </div>
//               <div className='col-sm-4 mb-2'>
//                 <label className='col-sm-4'>Assignee:</label>
//                 <PeoplePicker
//                   peoplePickerCntrlclassName='form-input'
//                   context={this.props.spfxContext as any}
//                   personSelectionLimit={4}
//                   showtooltip={false}
//                   required={false}
//                   disabled={false}
//                   ensureUser={true}
//                   principalTypes={[PrincipalType.User]}
//                   resolveDelay={500}
//                   onChange={(people: any) => this.setState({ selectedPeople: people })}
//                 />
//               </div>
//             </div>
//             <div className='col-sm-4 mb-2'>
//               <input className='btn btn-primary col-sm-2' type="submit" onClick={this.savedata} value="Submit" />
//               <input className='btn btn-secondary col-sm-2' type="button" onClick={this.updatedata} value="Update" />
//               <input className='btn btn-danger col-sm-2' type="button" onClick={() => this.deletedata()} value="Delete" />
//               <input className='btn btn-info col-sm-2 ml-2' type="button" onClick={this.readdata} value="Read Data" />
//             </div>
//           </form>
//           <div>
//             <h2>Fetched Items</h2>
//             <table className='table table-bordered'>
//               <thead>
//                 <tr>
//                   <th>Sr. No.</th>
//                   <th>Name</th>
//                   <th>Phone</th>
//                   <th>Email</th>
//                   <th>Message</th>
//                   <th>Date of Birth</th>
//                   <th>Gender</th>
//                   <th>Edit/Delete</th>
//                 </tr>
//               </thead>
//               <tbody>
//                 {this.state.items.map((item: any, index: number) => (
//                   <tr key={index}>
//                     <td>{index + 1}</td>
//                     <td>{item.Name}</td>
//                     <td>{item.MobileNumber}</td>
//                     <td>{item.Email}</td>
//                     <td>{item.Message}</td>
//                     <td>{item.DateofBirth}</td>
//                     <td>{item.Gender}</td>
//                     <td>
//                       <DefaultButton className='btn btn-primary m-1' text="Edit" onClick={() => this.editItem(item)} />
//                       <DefaultButton className='btn btn-danger m-1' text="Delete" onClick={() => this.deletedata(item.Id)} />
//                     </td>
//                   </tr>
//                 ))}
//               </tbody>
//             </table>
//           </div>
//         </div>
//       </>
//     );
//   }

//   private async savedata(event: React.FormEvent) {
//     event.preventDefault();

//     const obj = {
//       Name: (document.getElementById("name") as HTMLInputElement).value,
//       MobileNumber: (document.getElementById("phone") as HTMLInputElement).value,
//       Email: (document.getElementById("email") as HTMLInputElement).value,
//       Gender: (document.getElementById("gender") as HTMLInputElement).value,
//       DateofBirth: (document.getElementById("dob") as HTMLInputElement).value,
//       Message: (document.getElementById("message") as HTMLInputElement).value,
//       Assignee: this.state.selectedPeople.map((person: any) => person.id)
//     };

//     try {
//       const response = await fetch(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('AppointmentForm')/items`, {
//         method: 'POST',
//         headers: {
//           'Accept': 'application/json;odata=verbose',
//           'Content-Type': 'application/json;odata=verbose',
//           'X-RequestDigest': (document.getElementById('__REQUESTDIGEST') as HTMLInputElement).value
//         },
//         body: JSON.stringify(obj)
//       });
//       await response.json();
//       console.log("Data saved successfully.", );
//       alert("Data saved successfully.");
//       await this.readdata();
//       this.resetForm();
//     } catch (error) {
//       console.error("Error saving data", error);
//     }
//   }

//   private resetForm() {
//     (document.getElementById("name") as HTMLInputElement).value = '';
//     (document.getElementById("phone") as HTMLInputElement).value = '';
//     (document.getElementById("email") as HTMLInputElement).value = '';
//     (document.getElementById("gender") as HTMLInputElement).value = '';
//     (document.getElementById("dob") as HTMLInputElement).value = '';
//     (document.getElementById("message") as HTMLInputElement).value = '';
//     this.setState({ selectedPeople: [] });
//   }

//   private async updatedata() {
//     if (this.state.selectedItemID) {
//       const obj = {
//         Name: (document.getElementById("name") as HTMLInputElement).value,
//         MobileNumber: (document.getElementById("phone") as HTMLInputElement).value,
//         Email: (document.getElementById("email") as HTMLInputElement).value,
//         Gender: (document.getElementById("gender") as HTMLInputElement).value,
//         DateofBirth: (document.getElementById("dob") as HTMLInputElement).value,
//         Message: (document.getElementById("message") as HTMLInputElement).value
//       };

//       try {
//         const response = await fetch(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('AppointmentForm')/items(${this.state.selectedItemID})`, {
//           method: 'MERGE',
//           headers: {
//             'Accept': 'application/json;odata=verbose',
//             'Content-Type': 'application/json;odata=verbose',
//             'X-HTTP-Method': 'MERGE',
//             'IF-MATCH': '*',
//             'X-RequestDigest': (document.getElementById('__REQUESTDIGEST') as HTMLInputElement).value
//           },
//           body: JSON.stringify(obj)
//         });
//         if (response.ok) {
//           console.log("Data updated successfully.");
//           alert("Data updated successfully.");
//           await this.readdata();
//           this.resetForm();
//         } else {
//           console.error("Error updating data", response.statusText);
//         }
//       } catch (error) {
//         console.error("Error updating data", error);
//       }
//     } else {
//       console.error("No item selected for update");
//     }
//   }

//   private async deletedata(id?: number) {
//     const itemId = id || this.state.selectedItemID;
//     if (itemId) {
//       const confirmDeletion = confirm("Are you sure you want to delete this item?");
//       if (confirmDeletion) {
//         try {
//           const response = await fetch(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('AppointmentForm')/items(${itemId})`, {
//             method: 'DELETE',
//             headers: {
//               'Accept': 'application/json;odata=verbose',
//               'X-HTTP-Method': 'DELETE',
//               'IF-MATCH': '*',
//               'X-RequestDigest': (document.getElementById('__REQUESTDIGEST') as HTMLInputElement).value
//             }
//           });
//           if (response.ok) {
//             console.log("Data deleted successfully.");
//             alert("Data deleted successfully.");
//             await this.readdata();
//             this.setState({ selectedItemID: null });
//           } else {
//             console.error("Error deleting data", response.statusText);
//           }
//         } catch (error) {
//           console.error("Error deleting data", error);
//         }
//       } else {
//         console.log("Deletion cancelled by user.");
//       }
//     } else {
//       console.error("No item selected for deletion");
//     }
//   }

//   private async readdata() {
//     try {
//       const response = await fetch(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('AppointmentForm')/items`, {
//         method: 'GET',
//         headers: {
//           'Accept': 'application/json;odata=verbose',
//           'Content-Type': 'application/json;odata=verbose'
//         }
//       });
//       const result = await response.json();
//       this.setState({ items: result.d.results });
//       console.log("Data read successfully",);
//     } catch (error) {
//       console.error("Error reading data", error);
//     }
//   }

//   private editItem(item: any) {
//     this.setState({ selectedItemID: item.Id });
//     (document.getElementById("name") as HTMLInputElement).value = item.Name;
//     (document.getElementById("phone") as HTMLInputElement).value = item.MobileNumber;
//     (document.getElementById("email") as HTMLInputElement).value = item.Email;
//     (document.getElementById("gender") as HTMLInputElement).value = item.Gender;
//     (document.getElementById("dob") as HTMLInputElement).value = item.DateofBirth;
//     (document.getElementById("message") as HTMLInputElement).value = item.Message;
//   }
// }