package com.spring.upload.service.serviceImpl;

import com.spring.upload.model.Upload;
import com.spring.upload.repository.UploadRepository;
import com.spring.upload.service.UploadService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UploadServiceImpl implements UploadService {

    @Autowired
    UploadRepository uploadRepository;

    @Override
    public List<Upload> findAll() {
        return uploadRepository.findAll();
    }

    @Override
    public Upload findById(long id) {
        return uploadRepository.findById(id).get();
    }

    @Override
    public Upload save(Upload upload) {
        return uploadRepository.save(upload);
    }
}
