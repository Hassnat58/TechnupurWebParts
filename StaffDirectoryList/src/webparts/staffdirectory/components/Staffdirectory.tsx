import * as React from "react";
import styles from "./Staffdirectory.module.scss";
import type { IStaffdirectoryProps } from "./IStaffdirectoryProps";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp/presets/all";

interface StaffMember {

  EmployeeName: string;
  Title: string;
  Department: string;
  Phone: number;
  Email: string;
  Location: string;
  
}

const StaffDirectory: React.FC<IStaffdirectoryProps> = (props) => {
  const { context } = props;
  const [searchQuery, setSearchQuery] = useState("");
  const [filteredStaffData, setFilteredStaffData] = useState<StaffMember[]>([]);
  const [loader, setLoader] = useState(false);
  const [staffData, setStaffData] = useState<StaffMember[]>([]);

  useEffect(() => {
    sp.setup({ sp: { baseUrl: context.pageContext.web.absoluteUrl } });
    fetchData();
  }, []);

  const fetchData = async () => {
    setLoader(true);
    try {
      const data = await sp.web.lists
        .getByTitle("OrgChart")
        .items.select(
          "EmployeeName/Title", // Expand EmployeeName to get the Title (Full Name)
          "EmployeeName/EMail", // Fetch email from Person field
          "Department", // Choice field
          "Title",
          "Phone",
          "Location"
        )
        .expand("EmployeeName")(); // Expand the EmployeeName (Person field)
  
      let staffMembers: StaffMember[] = data.map((item: any) => ({
        EmployeeName: item.EmployeeName?.Title || "N/A", // Get the name from the Person field
        Email: item.EmployeeName?.EMail || "N/A", // Get Email from the Person field
        Title: item.Title,
        Department: item.Department || "N/A", // Ensure department is retrieved
        Phone: item.Phone ? item.Phone.toString() : "N/A", // Convert number to string
        Location: item.Location || "N/A",
      }));
  
      // Sorting the staffMembers array by EmployeeName in ascending order
      staffMembers = staffMembers.sort((a, b) => a.EmployeeName.localeCompare(b.EmployeeName));
  
      setStaffData(staffMembers);
      setFilteredStaffData(staffMembers);
    } catch (error) {
      console.error("Error fetching data:", error);
    } finally {
      setLoader(false);
    }
  };

  const handleSearch = (e: React.ChangeEvent<HTMLInputElement>) => {
    const query = e.target.value.toLowerCase();
    setSearchQuery(query);
    setLoader(true);
    setTimeout(() => {
      const filtered = staffData.filter((staff) =>
        Object.values(staff).some((value: string) => value.toLowerCase().includes(query))
      );
      setFilteredStaffData(filtered);
      setLoader(false);
    }, 500);
  };

  return (
    
    <div style={{ width: "100%" }}>
  <div className={styles.staffDirectoryContainer}>
    <div className={styles.header}>
      <div className={styles.titleDiv}>
        <svg
          width="24"
          height="16"
          viewBox="0 0 24 16"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          <path
            d="M0.947266 15.9167V13.5082C0.947266 12.9499 1.09135 12.4517 1.37952 12.0135C1.66768 11.5753 2.05254 11.2344 2.53408 10.9908C3.56361 10.4872 4.59882 10.0989 5.63972 9.82594C6.68081 9.55312 7.8249 9.41671 9.072 9.41671C10.3193 9.41671 11.4634 9.55312 12.5043 9.82594C13.5454 10.0989 14.5807 10.4872 15.6102 10.9908C16.0917 11.2344 16.4766 11.5753 16.7647 12.0135C17.0529 12.4517 17.197 12.9499 17.197 13.5082V15.9167H0.947266ZM19.3637 15.9167V13.3749C19.3637 12.6639 19.1896 11.9862 18.8415 11.3418C18.4932 10.6976 17.9992 10.1448 17.3595 9.68348C18.0859 9.79181 18.7755 9.95946 19.4284 10.1864C20.0811 10.4136 20.704 10.682 21.2971 10.9916C21.8569 11.2902 22.2891 11.6422 22.5939 12.0476C22.8987 12.4528 23.0511 12.8952 23.0511 13.3749V15.9167H19.3637ZM9.072 7.66659C8.02929 7.66659 7.13671 7.29536 6.39427 6.55292C5.65182 5.81029 5.2806 4.91763 5.2806 3.87492C5.2806 2.83221 5.65182 1.93963 6.39427 1.19719C7.13671 0.454565 8.02929 0.083252 9.072 0.083252C10.1147 0.083252 11.0074 0.454565 11.75 1.19719C12.4924 1.93963 12.8637 2.83221 12.8637 3.87492C12.8637 4.91763 12.4924 5.81029 11.75 6.55292C11.0074 7.29536 10.1147 7.66659 9.072 7.66659ZM18.426 3.87492C18.426 4.91763 18.0548 5.81029 17.3124 6.55292C16.5699 7.29536 15.6774 7.66659 14.6346 7.66659C14.5124 7.66659 14.3569 7.65268 14.168 7.62488C13.979 7.59707 13.8234 7.56656 13.7013 7.53334C14.1285 7.01965 14.4569 6.44982 14.6864 5.82384C14.9157 5.19785 15.0303 4.54785 15.0303 3.87384C15.0303 3.19964 14.9133 2.55208 14.6793 1.93115C14.4453 1.3104 14.1193 0.738939 13.7013 0.216773C13.8568 0.161161 14.0124 0.12505 14.168 0.108439C14.3235 0.0916476 14.479 0.083252 14.6346 0.083252C15.6774 0.083252 16.5699 0.454565 17.3124 1.19719C18.0548 1.93963 18.426 2.83221 18.426 3.87492ZM2.57199 14.2917H15.572V13.5082C15.572 13.282 15.5154 13.0806 15.4022 12.9042C15.2892 12.7278 15.1097 12.5736 14.8638 12.4416C13.9722 11.982 13.0541 11.6337 12.1097 11.3968C11.1652 11.1601 10.1526 11.0417 9.072 11.0417C7.99155 11.0417 6.97909 11.1601 6.0346 11.3968C5.09011 11.6337 4.17208 11.982 3.28049 12.4416C3.03458 12.5736 2.85502 12.7278 2.74181 12.9042C2.6286 13.0806 2.57199 13.282 2.57199 13.5082V14.2917ZM9.072 6.04159C9.66783 6.04159 10.1779 5.82943 10.6022 5.40513C11.0265 4.98082 11.2387 4.47075 11.2387 3.87492C11.2387 3.27909 11.0265 2.76902 10.6022 2.34471C10.1779 1.9204 9.66783 1.70825 9.072 1.70825C8.47616 1.70825 7.96609 1.9204 7.54179 2.34471C7.11748 2.76902 6.90533 3.27909 6.90533 3.87492C6.90533 4.47075 7.11748 4.98082 7.54179 5.40513C7.96609 5.82943 8.47616 6.04159 9.072 6.04159Z"
            fill="white"
          />
        </svg>
        <h3>&nbsp; Staff Directory</h3>
      </div>
    </div>
    <div className={styles.searchBar}>
      <input
        type="text"
        placeholder="Search"
        value={searchQuery}
        onChange={(e) => {
          handleSearch(e);
        }}
      />
      <button className={styles.searchButton}>
        <svg
          width="17"
          height="17"
          viewBox="0 0 17 17"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          <path
            d="M15.6777 17L9.72766 11.05C9.25543 11.4278 8.71238 11.7269 8.09849 11.9472C7.4846 12.1676 6.83136 12.2778 6.13877 12.2778C4.42303 12.2778 2.97094 11.6836 1.78252 10.4951C0.594091 9.30671 -0.00012207 7.85463 -0.00012207 6.13889C-0.00012207 4.42315 0.594091 2.97106 1.78252 1.78264C2.97094 0.594213 4.42303 0 6.13877 0C7.85451 0 9.30659 0.594213 10.495 1.78264C11.6834 2.97106 12.2777 4.42315 12.2777 6.13889C12.2777 6.83148 12.1675 7.48472 11.9471 8.09861C11.7267 8.7125 11.4277 9.25556 11.0499 9.72778L16.9999 15.6778L15.6777 17ZM6.13877 10.3889C7.31932 10.3889 8.32279 9.97569 9.14918 9.14931C9.97557 8.32292 10.3888 7.31944 10.3888 6.13889C10.3888 4.95833 9.97557 3.95486 9.14918 3.12847C8.32279 2.30208 7.31932 1.88889 6.13877 1.88889C4.95821 1.88889 3.95474 2.30208 3.12835 3.12847C2.30196 3.95486 1.88877 4.95833 1.88877 6.13889C1.88877 7.31944 2.30196 8.32292 3.12835 9.14931C3.95474 9.97569 4.95821 10.3889 6.13877 10.3889Z"
            fill="#005A9C"
          />
        </svg>
      </button>
    </div>
    <div className={styles.tableContainer}>
      <div className={styles.table}>
        {/* Table Header */}
        <div className={styles.tableHeader}>
        
          <div>Name</div>
          <div>Designation</div>
          <div>Department</div>
          <div>Phone No.</div>
          <div>Email</div>
          <div>Location</div>
          
        </div>

        {/* Table Body */}
        <div className={styles.tableBody}>
          {!loader && filteredStaffData?.length > 0 ? (
            filteredStaffData.map((staff: any, index: number) => (
              <div className={styles.tableRow} key={index}>
                
                    <div>{staff.EmployeeName}</div>
                    <div>{staff.Title}</div>
                    <div>{staff.Department}</div>
                    <div>{staff.Phone}</div>
                    <div>{staff.Email}</div>
                    <div>{staff.Location}</div>
               
              </div>
            ))
          ) : loader ? (
            <div
              style={{ textAlign: "center", display: "block" }}
              className={styles.tableRow}
            >
              <b>Loading Staff Directory...</b>
            </div>
          ) : (
            <div
              style={{
                textAlign: "center",
                display: "block",
                marginTop: "10px",
              }}
              className={styles.tableRow}
            >
              No Staff Directory Found!
            </div>
          )}
        </div>
      </div>
    </div>
  </div>
  </div>

    
  );
};

export default StaffDirectory;
