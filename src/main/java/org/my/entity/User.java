package org.my.entity;

import org.my.annotation.ThisIgnore;
import org.my.annotation.TimePattern;

public class User {
	
	public static int test;
	private Integer id;
	private Integer status;
	@TimePattern
	private String createtime;
	private String username;
	private String password;
	public static int getTest() {
		return test;
	}
	public static void setTest(int test) {
		User.test = test;
	}

	@ThisIgnore
	private String ignore;
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public Integer getStatus() {
		return status;
	}
	public void setStatus(Integer status) {
		this.status = status;
	}
	public String getCreatetime() {
		return createtime;
	}
	public void setCreatetime(String createtime) {
		this.createtime = createtime;
	}
	public String getUsername() {
		return username;
	}
	public void setUsername(String username) {
		this.username = username;
	}
	public String getPassword() {
		return password;
	}
	public void setPassword(String password) {
		this.password = password;
	}
	public String getIgnore() {
		return ignore;
	}
	public void setIgnore(String ignore) {
		this.ignore = ignore;
	}
	@Override
	public String toString() {
		return "id:" + id + ";status:" + status + ";createtime:" + createtime + ";username:" + username + ";password:" + password;
	}
}
