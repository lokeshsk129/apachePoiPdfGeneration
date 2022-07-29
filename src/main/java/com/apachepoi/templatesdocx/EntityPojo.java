package com.apachepoi.templatesdocx;

import java.util.Map;

public class EntityPojo {
	
	private String firstName;
	private String lastName;
	private String gender;
	private String mobilePhone;
	private String email;
	private String homeAddress;	
	private String dateOfBirth;
	private Map<String, String>textInfo;
	
	
	private String imagePath;
	private Map<String, String>mediaInfo;
	
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	public String getLastName() {
		return lastName;
	}
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	public String getGender() {
		return gender;
	}
	public void setGender(String gender) {
		this.gender = gender;
	}
	public String getMobilePhone() {
		return mobilePhone;
	}
	public void setMobilePhone(String mobilePhone) {
		this.mobilePhone = mobilePhone;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getHomeAddress() {
		return homeAddress;
	}
	public void setHomeAddress(String homeAddress) {
		this.homeAddress = homeAddress;
	}
	public String getDateOfBirth() {
		return dateOfBirth;
	}
	public void setDateOfBirth(String dateOfBirth) {
		this.dateOfBirth = dateOfBirth;
	}
	public Map<String, String> getTextInfo() {
		return textInfo;
	}
	public void setTextInfo(Map<String, String> textInfo) {
		this.textInfo = textInfo;
	}
	public String getImagePath() {
		return imagePath;
	}
	public void setImagePath(String imagePath) {
		this.imagePath = imagePath;
	}
	public Map<String, String> getMediaInfo() {
		return mediaInfo;
	}
	public void setMediaInfo(Map<String, String> mediaInfo) {
		this.mediaInfo = mediaInfo;
	}
	public EntityPojo(String firstName, String lastName, String gender, String mobilePhone, String email,
			String homeAddress, String dateOfBirth, Map<String, String> textInfo, String imagePath,
			Map<String, String> mediaInfo) {
		super();
		this.firstName = firstName;
		this.lastName = lastName;
		this.gender = gender;
		this.mobilePhone = mobilePhone;
		this.email = email;
		this.homeAddress = homeAddress;
		this.dateOfBirth = dateOfBirth;
		this.textInfo = textInfo;
		this.imagePath = imagePath;
		this.mediaInfo = mediaInfo;
	}
	public EntityPojo() {
		super();
		
	}
	@Override
	public String toString() {
		return "EntityPojo [firstName=" + firstName + ", lastName=" + lastName + ", gender=" + gender + ", mobilePhone="
				+ mobilePhone + ", email=" + email + ", homeAddress=" + homeAddress + ", dateOfBirth=" + dateOfBirth
				+ ", textInfo=" + textInfo + ", imagePath=" + imagePath + ", mediaInfo=" + mediaInfo + "]";
	}
	
	

}
