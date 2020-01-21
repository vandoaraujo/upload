package com.spring.upload.service;

import com.spring.upload.model.Upload;

import java.util.List;

public interface UploadService {

    List<Upload> findAll();

    Upload findById(long id);

    Upload save(Upload upload);

}
